"""Impressão de verdade no Windows, numa impressora que grava em arquivo.

No CI é criada a impressora ``CI-Printer`` com o driver "Microsoft Print to PDF" e
porta apontando para um arquivo; aqui conferimos que o PDF gerado tem as páginas
esperadas. Localmente:

    $env:PRINTER_CAV_TEST_PRINTER = "Nome da impressora"
    $env:PRINTER_CAV_TEST_OUTPUT  = "C:\\caminho\\saida.pdf"   # opcional
    pytest -m windows
"""

from __future__ import annotations

import os
import queue
import time
from pathlib import Path

import pytest
from PIL import Image

PRINTER = os.environ.get("PRINTER_CAV_TEST_PRINTER")
OUTPUT = (
    Path(os.environ["PRINTER_CAV_TEST_OUTPUT"])
    if os.environ.get("PRINTER_CAV_TEST_OUTPUT")
    else None
)
OUTPUT_IS_PDF = os.environ.get("PRINTER_CAV_TEST_FORMAT", "pdf") == "pdf"

pytestmark = [
    pytest.mark.windows,
    pytest.mark.skipif(
        os.name != "nt" or not PRINTER, reason="requer Windows e impressora de teste"
    ),
]


def _wait_spool_and_output(timeout: float = 120) -> None:
    """Espera a fila esvaziar e o arquivo de saída parar de crescer."""
    from printer_app.printers import pending_jobs

    deadline = time.monotonic() + timeout
    last = -1
    while time.monotonic() < deadline:
        if pending_jobs(PRINTER) == 0:
            if OUTPUT is None:
                time.sleep(1.0)
                return
            size = OUTPUT.stat().st_size if OUTPUT.exists() else -1
            if size > 0 and size == last:
                return
            last = size
        time.sleep(0.5)
    raise AssertionError("o trabalho não saiu da fila a tempo")


@pytest.fixture
def output():
    if OUTPUT is not None and OUTPUT.exists():
        OUTPUT.unlink()
    yield OUTPUT


def _printed_pages() -> int | None:
    if OUTPUT is None or not OUTPUT_IS_PDF:
        return None
    from printer_app.printing import pdf

    return pdf.page_count(OUTPUT)


def _entry(path: Path):
    from printer_app.files import Entry, kind_of

    return Entry(path, path.name, "teste", kind_of(path))


def test_printer_is_listed_and_readable():
    from printer_app import printers

    names = [p.name for p in printers.list_printers()]
    assert PRINTER in names
    printers.printer_problem(PRINTER)  # não pode exigir admin nem lançar exceção


def test_print_pdf_goes_to_selected_printer(output, sample_pdf):
    from printer_app.printing import PrintContext, print_entry

    assert print_entry(_entry(sample_pdf(pages=3)), PrintContext(PRINTER)) == 3
    _wait_spool_and_output()
    if OUTPUT_IS_PDF and OUTPUT is not None:
        assert _printed_pages() == 3


def test_print_image_with_copies(output, tmp_path):
    from printer_app.printing import PrintContext, print_entry

    png = tmp_path / "comprovante.png"
    Image.new("RGBA", (1200, 800), (47, 93, 168, 255)).save(png)
    assert print_entry(_entry(png), PrintContext(PRINTER, copies=2)) == 2
    _wait_spool_and_output()
    if OUTPUT_IS_PDF and OUTPUT is not None:
        assert _printed_pages() == 2


def test_print_text_document_via_office_or_libreoffice(output, tmp_path):
    from printer_app.printing import PrintContext, print_entry
    from printer_app.printing.office import OfficeSession, find_libreoffice

    with OfficeSession() as session:
        if session.word() is None and find_libreoffice() is None:
            pytest.skip("sem Microsoft Office nem LibreOffice")
        txt = tmp_path / "comunicado.txt"
        txt.write_text("Comunicado às consultoras\n" * 5, encoding="utf-8")
        print_entry(_entry(txt), PrintContext(PRINTER, office=session))
    _wait_spool_and_output()
    if OUTPUT_IS_PDF and OUTPUT is not None:
        assert _printed_pages() >= 1


def test_batch_with_real_backend_isolates_errors(output, sample_pdf, tmp_path):
    from printer_app.backend import WindowsBackend
    from printer_app.jobs import PrintJob, Status

    bad = tmp_path / "corrompido.pdf"
    bad.write_bytes(b"%PDF-1.4 lixo")
    backend = WindowsBackend()
    events: queue.Queue = queue.Queue()
    items = [(0, _entry(bad)), (1, _entry(sample_pdf("ok.pdf", pages=1)))]
    job = PrintJob(items, PRINTER, 1, events, backend.print_fn, backend.session_factory)
    job.start()
    job.join(120)
    final = {}
    while not events.empty():
        event = events.get()
        if event.index is not None:
            final[event.index] = event.status
    assert final == {0: Status.ERROR, 1: Status.DONE}
    _wait_spool_and_output()


@pytest.mark.skipif(OUTPUT is None or not OUTPUT_IS_PDF, reason="requer saída em PDF")
def test_vector_pdf_is_much_lighter_than_raster(output, tmp_path):
    """PDF de texto impresso em vetor gera menos dados que o mesmo PDF em bitmap.

    O ganho de tempo é medido pelo ``tools/benchmark_print.py`` no CI.
    """
    from pdfgen import make_text_pdf

    from printer_app.printing import gdi, pdf

    doc = make_text_pdf(tmp_path / "texto.pdf", pages=5)
    sizes = {}
    for mode in ("raster", "vetor"):
        if OUTPUT.exists():
            OUTPUT.unlink()
        if mode == "raster":
            gdi.print_pages(PRINTER, "raster", lambda dpi: pdf.pdf_pages(doc, dpi))
        else:
            gdi.print_pdf_vector(PRINTER, doc, "vetor")
        _wait_spool_and_output()
        assert _printed_pages() == 5
        sizes[mode] = OUTPUT.stat().st_size
    print(f"\nsaída raster: {sizes['raster']:,} bytes | vetor: {sizes['vetor']:,} bytes")
    assert sizes["vetor"] < sizes["raster"]
