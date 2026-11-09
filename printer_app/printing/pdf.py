"""Renderização de PDF com PDFium (o mesmo motor do Chrome) — sem leitor de PDF externo."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from PIL import Image

MAX_RENDER_DPI = 300


class PdfError(Exception):
    """PDF ilegível, com mensagem pronta para o usuário."""


def _open(path: Path):
    import pypdfium2 as pdfium

    try:
        return pdfium.PdfDocument(str(path))
    except pdfium.PdfiumError as exc:
        if "password" in str(exc).lower():
            raise PdfError("PDF protegido por senha") from exc
        raise PdfError(f"PDF corrompido ou ilegível ({exc})") from exc


def page_count(path: Path) -> int:
    return inspect(path)[0]


def inspect(path: Path) -> tuple[int, bool]:
    """Valida o PDF e devolve ``(páginas, precisa_raster)``.

    Formulários XFA só são desenhados corretamente pela renderização completa em bitmap;
    todo o resto pode ir para a impressora em vetor.
    """
    import pypdfium2.raw as pdfium_c

    doc = _open(path)
    try:
        pages = len(doc)
        if pages == 0:
            raise PdfError("PDF sem páginas")
        xfa = doc.get_formtype() in (pdfium_c.FORMTYPE_XFA_FULL, pdfium_c.FORMTYPE_XFA_FOREGROUND)
        return pages, xfa
    finally:
        doc.close()


def pdf_pages(path: Path, dpi: int = MAX_RENDER_DPI) -> Iterator[Image.Image]:
    """Gera as páginas renderizadas uma a uma (inclui campos de formulário preenchidos)."""
    scale = min(dpi, MAX_RENDER_DPI) / 72
    doc = _open(path)
    try:
        doc.init_forms()
        for index in range(len(doc)):
            page = doc[index]
            try:
                bitmap = page.render(scale=scale, may_draw_forms=True)
                yield bitmap.to_pil().convert("RGB")
            finally:
                page.close()
    finally:
        doc.close()
