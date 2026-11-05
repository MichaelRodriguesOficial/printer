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
    doc = _open(path)
    try:
        return len(doc)
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
