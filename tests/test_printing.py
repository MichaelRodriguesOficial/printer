from __future__ import annotations

import threading
from pathlib import Path

import pytest
from PIL import Image

from printer_app.files import Entry, FileKind
from printer_app.printing import PrintContext, PrintError, gdi, images, pdf, print_entry
from printer_app.printing.gdi import Cancelled, fit_box, orient_for_page

# --- imagens -------------------------------------------------------------------


def test_transparent_png_gets_white_background(tmp_path):
    path = tmp_path / "logo.png"
    img = Image.new("RGBA", (10, 10), (0, 0, 0, 0))
    img.putpixel((5, 5), (255, 0, 0, 255))
    img.save(path)
    (page,) = list(images.image_pages(path))
    assert page.mode == "RGB"
    assert page.getpixel((0, 0)) == (255, 255, 255)  # na v1 saía preto
    assert page.getpixel((5, 5)) == (255, 0, 0)


def test_palette_transparency_and_gif_first_frame(tmp_path):
    path = tmp_path / "anim.gif"
    frames = [Image.new("P", (8, 8), i) for i in range(3)]
    frames[0].save(path, save_all=True, append_images=frames[1:], transparency=0)
    pages = list(images.image_pages(path))
    assert len(pages) == 1
    assert pages[0].getpixel((0, 0)) == (255, 255, 255)


def test_exif_orientation_is_applied(tmp_path):
    path = tmp_path / "celular.jpg"
    exif = Image.Exif()
    exif[0x0112] = 6  # rotacionar 90° (foto de celular "em pé")
    Image.new("RGB", (40, 20), "blue").save(path, exif=exif)
    (page,) = list(images.image_pages(path))
    assert page.size == (20, 40)


def test_multipage_tiff_prints_all_pages(tmp_path):
    path = tmp_path / "scan.tif"
    frames = [Image.new("RGB", (30, 40), c) for c in ("red", "green", "blue")]
    frames[0].save(path, save_all=True, append_images=frames[1:])
    assert images.count_pages(path) == 3
    assert [p.getpixel((0, 0)) for p in images.image_pages(path)] == [
        (255, 0, 0),
        (0, 128, 0),
        (0, 0, 255),
    ]


# --- PDF -----------------------------------------------------------------------


def test_pdf_render_pages(sample_pdf):
    path = sample_pdf(pages=3)
    assert pdf.page_count(path) == 3
    pages = list(pdf.pdf_pages(path, dpi=144))
    assert len(pages) == 3
    assert pages[0].mode == "RGB"
    assert pages[0].size == (1190, 1684)  # A4 a 144 dpi


def test_pdf_render_dpi_is_capped(sample_pdf):
    page = next(pdf.pdf_pages(sample_pdf(pages=1), dpi=1200))
    assert abs(page.width - 595 * pdf.MAX_RENDER_DPI / 72) <= 1


def test_pdf_password_and_corrupted(tmp_path, sample_pdf):
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    for page in PdfReader(sample_pdf(pages=1)).pages:
        writer.add_page(page)
    writer.encrypt("1234")
    locked = tmp_path / "senha.pdf"
    with open(locked, "wb") as fh:
        writer.write(fh)
    with pytest.raises(pdf.PdfError, match="senha"):
        pdf.page_count(locked)

    bad = tmp_path / "ruim.pdf"
    bad.write_bytes(b"%PDF-1.4 lixo")
    with pytest.raises(pdf.PdfError, match="corrompido"):
        list(pdf.pdf_pages(bad))


# --- GDI -------------------------------------------------------------------------


def test_fit_box_and_orientation():
    assert fit_box(100, 200, 1000, 1000) == (250, 0, 750, 1000)
    landscape = Image.new("RGB", (300, 100))
    assert orient_for_page(landscape, 2400, 3400).size == (100, 300)
    assert orient_for_page(landscape, 3400, 2400).size == (300, 100)
    square = Image.new("RGB", (100, 100))
    assert orient_for_page(square, 2400, 3400) is square


def test_print_pages_copies_and_target_printer(fake_gdi):
    seen_dpi = []

    def pages(dpi):
        seen_dpi.append(dpi)
        return [Image.new("RGB", (10, 20)), Image.new("RGB", (10, 20))]

    printed = gdi.print_pages("BRA3-ATENDIMENTO", "doc.pdf", pages, copies=2)
    calls = fake_gdi["calls"]
    assert printed == 4
    assert calls[0] == ("CreatePrinterDC", "BRA3-ATENDIMENTO")  # impressora escolhida
    assert calls[1] == ("StartDoc", "doc.pdf")
    assert calls.count(("StartPage",)) == 4
    assert calls[-2:] == [("EndDoc",), ("DeleteDC",)]
    assert seen_dpi == [300, 300]
    assert fake_gdi["draws"][0] == ((10, 20), (350, 0, 2050, 3400))  # centralizado


def test_print_pages_aborts_on_error_and_cancel(fake_gdi):
    fake_gdi["fail_on_page"] = 2
    with pytest.raises(RuntimeError):
        gdi.print_pages("P", "x", lambda dpi: [Image.new("RGB", (5, 5))] * 3)
    assert fake_gdi["calls"][-2:] == [("AbortDoc",), ("DeleteDC",)]

    fake_gdi["calls"].clear()
    fake_gdi["fail_on_page"] = None
    cancel = threading.Event()

    def pages(dpi):
        yield Image.new("RGB", (5, 5))
        cancel.set()
        yield Image.new("RGB", (5, 5))

    with pytest.raises(Cancelled):
        gdi.print_pages("P", "x", pages, cancel=cancel)
    assert fake_gdi["calls"].count(("EndPage",)) == 1
    assert fake_gdi["calls"][-2:] == [("AbortDoc",), ("DeleteDC",)]


def test_print_pages_empty_document(fake_gdi):
    with pytest.raises(ValueError, match="sem páginas"):
        gdi.print_pages("P", "x", lambda dpi: [])
    assert ("AbortDoc",) in fake_gdi["calls"]


# --- despacho ------------------------------------------------------------------


def _entry(path: Path, kind: FileKind, error: str | None = None) -> Entry:
    return Entry(path, path.name, "pasta", kind, error)


def test_print_entry_pdf_and_image_end_to_end_with_fake_gdi(fake_gdi, sample_pdf, tmp_path):
    ctx = PrintContext("Financeiro", copies=1)
    assert print_entry(_entry(sample_pdf(pages=2), FileKind.PDF), ctx) == 2
    png = tmp_path / "a.png"
    Image.new("RGB", (50, 50)).save(png)
    assert print_entry(_entry(png, FileKind.IMAGE), ctx) == 1
    printers = [c[1] for c in fake_gdi["calls"] if c[0] == "CreatePrinterDC"]
    assert printers == ["Financeiro", "Financeiro"]


def test_print_entry_errors_become_print_error(tmp_path):
    bad = tmp_path / "ruim.pdf"
    bad.write_bytes(b"lixo")
    ctx = PrintContext("P")
    with pytest.raises(PrintError, match="corrompido"):
        print_entry(_entry(bad, FileKind.PDF), ctx)  # falha antes de abrir a impressora
    with pytest.raises(PrintError, match="Imagem corrompida"):
        print_entry(_entry(bad, FileKind.IMAGE), ctx)
    with pytest.raises(PrintError, match="não suportado"):
        print_entry(_entry(bad, FileKind.UNSUPPORTED), ctx)
    with pytest.raises(PrintError, match="senha"):
        print_entry(_entry(bad, FileKind.ARCHIVE, "Compactado protegido por senha"), ctx)
