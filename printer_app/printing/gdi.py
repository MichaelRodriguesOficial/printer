"""Envio de páginas diretamente para a impressora escolhida via GDI.

Dois caminhos, ambos sem depender do leitor de PDF padrão nem da impressora padrão do
Windows (o trabalho vai sempre para ``printer_name``, com as preferências salvas pelo
usuário em "Preferências da impressora": frente e verso, cor, bandeja...):

* **PDF vetorial** (:func:`print_pdf_vector`): o PDFium desenha a página direto no
  device context da impressora, como o Foxit e o Chrome fazem. Texto e vetores seguem
  como comandos de desenho, e o spool fica com KB por página em vez de dezenas de MB.
* **Raster** (:func:`print_pages`): imagens e PDFs que precisam de renderização
  completa (formulários XFA) seguem como bitmap, reduzido ao tamanho da página.
"""

from __future__ import annotations

import threading
from collections.abc import Callable, Iterable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from PIL import Image

# Índices de GetDeviceCaps (wingdi.h)
HORZRES, VERTRES, LOGPIXELSX = 8, 10, 88
# Flags de FPDF_RenderPage: anotações (campos preenchidos, carimbos) + modo impressão.
FPDF_ANNOT, FPDF_PRINTING = 0x01, 0x800

PageFactory = Callable[[int], Iterable[Image.Image]]


class Cancelled(Exception):  # noqa: N818 - sinal de cancelamento, não um erro
    """O usuário cancelou a impressão."""


@dataclass
class PrinterPage:
    """Área imprimível do papel atual, em pixels do dispositivo."""

    width: int
    height: int
    dpi: int


def fit_box(img_w: float, img_h: float, page_w: int, page_h: int) -> tuple[int, int, int, int]:
    """Retângulo (x0, y0, x1, y1) que encaixa o conteúdo centralizado na área imprimível."""
    scale = min(page_w / img_w, page_h / img_h)
    w, h = max(1, round(img_w * scale)), max(1, round(img_h * scale))
    x, y = (page_w - w) // 2, (page_h - h) // 2
    return x, y, x + w, y + h


def needs_rotation(width: float, height: float, page_w: int, page_h: int) -> bool:
    """Conteúdo paisagem em papel retrato (ou o contrário) deve ser girado 90°."""
    return width != height and (width > height) != (page_w > page_h)


def orient_for_page(img: Image.Image, page_w: int, page_h: int) -> Image.Image:
    """Gira páginas/imagens paisagem para aproveitar o papel retrato (e vice-versa)."""
    if needs_rotation(img.width, img.height, page_w, page_h):
        return img.rotate(90, expand=True)
    return img


@contextmanager
def printer_job(printer_name: str, doc_name: str) -> Iterator[tuple[object, PrinterPage]]:
    """Abre um trabalho na impressora; em caso de erro/cancelamento o trabalho é abortado."""
    import win32ui

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    try:
        info = PrinterPage(
            dc.GetDeviceCaps(HORZRES),
            dc.GetDeviceCaps(VERTRES),
            dc.GetDeviceCaps(LOGPIXELSX) or 300,
        )
        dc.StartDoc(doc_name[:120])
        try:
            yield dc, info
            dc.EndDoc()
        except BaseException:
            dc.AbortDoc()
            raise
    finally:
        dc.DeleteDC()


def _check_cancel(cancel: threading.Event | None) -> None:
    if cancel is not None and cancel.is_set():
        raise Cancelled


def print_pages(
    printer_name: str,
    doc_name: str,
    pages: PageFactory,
    copies: int = 1,
    cancel: threading.Event | None = None,
) -> int:
    """Imprime as páginas geradas por ``pages(dpi)`` e devolve quantas páginas saíram.

    ``pages`` é chamada uma vez por cópia, gerando as páginas sob demanda (cópias
    agrupadas, sem manter o documento inteiro renderizado na memória).
    """
    from PIL import ImageWin

    printed = 0
    with printer_job(printer_name, doc_name) as (dc, paper):
        for _ in range(max(1, copies)):
            for page in pages(paper.dpi):
                _check_cancel(cancel)
                page = orient_for_page(page, paper.width, paper.height)
                box = fit_box(page.width, page.height, paper.width, paper.height)
                page = _shrink_to_box(page, box)
                dc.StartPage()
                ImageWin.Dib(page).draw(dc.GetHandleOutput(), box)
                dc.EndPage()
                printed += 1
        if printed == 0:
            raise ValueError("Documento sem páginas")
    return printed


def _shrink_to_box(img: Image.Image, box: tuple[int, int, int, int]) -> Image.Image:
    """Não envia mais pixels do que a impressora vai usar (foto de 12 MP → tamanho do papel)."""
    w, h = box[2] - box[0], box[3] - box[1]
    if img.width <= w and img.height <= h:
        return img  # imagem pequena: o driver amplia, com menos dados no spool
    return img.resize((w, h), Image.Resampling.BILINEAR, reducing_gap=2.0)


def print_pdf_vector(
    printer_name: str,
    path: Path,
    doc_name: str,
    copies: int = 1,
    cancel: threading.Event | None = None,
) -> int:
    """Imprime um PDF desenhando cada página direto na impressora (vetores e texto)."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c

    doc = pdfium.PdfDocument(str(path))
    printed = 0
    try:
        with printer_job(printer_name, doc_name) as (dc, paper):
            hdc = dc.GetSafeHdc()
            for _ in range(max(1, copies)):
                for index in range(len(doc)):
                    _check_cancel(cancel)
                    page = doc[index]
                    try:
                        width, height = page.get_size()
                        rotate = needs_rotation(width, height, paper.width, paper.height)
                        if rotate:
                            width, height = height, width
                        x0, y0, x1, y1 = fit_box(width, height, paper.width, paper.height)
                        dc.StartPage()
                        pdfium_c.FPDF_RenderPage(
                            hdc, page.raw, x0, y0, x1 - x0, y1 - y0,
                            1 if rotate else 0, FPDF_ANNOT | FPDF_PRINTING,
                        )  # fmt: skip
                        dc.EndPage()
                    finally:
                        page.close()
                    printed += 1
            if printed == 0:
                raise ValueError("Documento sem páginas")
    finally:
        doc.close()
    return printed


def vector_available() -> bool:
    """``FPDF_RenderPage`` (desenho em HDC) só existe nos binários do PDFium para Windows."""
    import pypdfium2.raw as pdfium_c

    return hasattr(pdfium_c, "FPDF_RenderPage")
