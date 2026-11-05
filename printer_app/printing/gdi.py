"""Envio de páginas (imagens PIL) diretamente para a impressora escolhida via GDI.

Usado por imagens e PDFs: não depende do leitor de PDF padrão nem da impressora padrão
do Windows — o trabalho vai sempre para ``printer_name``, com as preferências salvas
pelo usuário em "Preferências da impressora" (frente e verso, cor, bandeja...).
"""

from __future__ import annotations

import threading
from collections.abc import Callable, Iterable

from PIL import Image

# Índices de GetDeviceCaps (wingdi.h)
HORZRES, VERTRES, LOGPIXELSX = 8, 10, 88

PageFactory = Callable[[int], Iterable[Image.Image]]


class Cancelled(Exception):  # noqa: N818 - sinal de cancelamento, não um erro
    """O usuário cancelou a impressão."""


def fit_box(img_w: int, img_h: int, page_w: int, page_h: int) -> tuple[int, int, int, int]:
    """Retângulo (x0, y0, x1, y1) que encaixa a imagem centralizada na área imprimível."""
    scale = min(page_w / img_w, page_h / img_h)
    w, h = max(1, round(img_w * scale)), max(1, round(img_h * scale))
    x, y = (page_w - w) // 2, (page_h - h) // 2
    return x, y, x + w, y + h


def orient_for_page(img: Image.Image, page_w: int, page_h: int) -> Image.Image:
    """Gira páginas/imagens paisagem para aproveitar o papel retrato (e vice-versa)."""
    if img.width != img.height and (img.width > img.height) != (page_w > page_h):
        return img.rotate(90, expand=True)
    return img


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
    import win32ui
    from PIL import ImageWin

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    printed = 0
    try:
        page_w, page_h = dc.GetDeviceCaps(HORZRES), dc.GetDeviceCaps(VERTRES)
        dpi = dc.GetDeviceCaps(LOGPIXELSX) or 300
        dc.StartDoc(doc_name[:120])
        try:
            for _ in range(max(1, copies)):
                for page in pages(dpi):
                    if cancel is not None and cancel.is_set():
                        raise Cancelled
                    page = orient_for_page(page, page_w, page_h)
                    dc.StartPage()
                    ImageWin.Dib(page).draw(
                        dc.GetHandleOutput(), fit_box(page.width, page.height, page_w, page_h)
                    )
                    dc.EndPage()
                    printed += 1
            if printed == 0:
                raise ValueError("Documento sem páginas")
            dc.EndDoc()
        except BaseException:
            dc.AbortDoc()
            raise
    finally:
        dc.DeleteDC()
    return printed
