"""Despacho de cada item da lista para o método de impressão adequado."""

from __future__ import annotations

import shutil
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path

from ..files import Entry, FileKind
from . import gdi, images, pdf
from .gdi import Cancelled
from .office import OfficeError, OfficeSession, _PrintViaPdf
from .pdf import PdfError

__all__ = ["PDF_MODES", "Cancelled", "PrintContext", "PrintError", "print_entry"]


class PrintError(Exception):
    """Falha de impressão com mensagem pronta para o usuário."""


PDF_MODES = ("vector", "raster")


@dataclass
class PrintContext:
    printer: str
    copies: int = 1
    cancel: threading.Event = field(default_factory=threading.Event)
    office: OfficeSession = field(default_factory=OfficeSession)
    pdf_mode: str = "vector"
    printer_dc: gdi.PrinterDC | None = None

    @property
    def target(self) -> str | gdi.PrinterDC:
        """DC reaproveitado no lote (se houver) ou o nome da impressora."""
        return self.printer_dc if self.printer_dc is not None else self.printer


def print_entry(entry: Entry, ctx: PrintContext) -> int | None:
    """Imprime um item e devolve o número de páginas enviadas (``None`` se desconhecido)."""
    if entry.error:
        raise PrintError(entry.error)
    try:
        if entry.kind is FileKind.PDF:
            return _print_pdf(entry.path, entry.display, ctx)
        if entry.kind is FileKind.IMAGE:
            _check_image(entry.path)
            return gdi.print_pages(
                ctx.target,
                entry.display,
                lambda _dpi: images.image_pages(entry.path),
                ctx.copies,
                ctx.cancel,
            )
        if entry.kind is FileKind.WORD:
            return _print_office(entry, ctx, ctx.office.print_word)
        if entry.kind is FileKind.EXCEL:
            return _print_office(entry, ctx, ctx.office.print_excel)
    except Cancelled:
        raise
    except (PdfError, OfficeError, PrintError) as exc:
        raise PrintError(str(exc)) from exc
    except OSError as exc:
        raise PrintError(_os_message(exc)) from exc
    raise PrintError("Tipo de arquivo não suportado")


def _check_image(path: Path) -> None:
    """Valida a imagem antes de abrir um trabalho na impressora."""
    from PIL import UnidentifiedImageError

    try:
        images.count_pages(path)
    except (UnidentifiedImageError, SyntaxError, ValueError) as exc:
        raise PrintError("Imagem corrompida ou formato não reconhecido") from exc


def _print_pdf(path: Path, title: str, ctx: PrintContext) -> int:
    # Valida (senha/corrompido) antes de abrir um trabalho na impressora.
    _pages, needs_raster = pdf.inspect(path)
    if ctx.pdf_mode == "vector" and not needs_raster and gdi.vector_available():
        return gdi.print_pdf_vector(ctx.target, path, title, ctx.copies, ctx.cancel)
    return gdi.print_pages(
        ctx.target, title, lambda dpi: pdf.pdf_pages(path, dpi), ctx.copies, ctx.cancel
    )


def _print_office(entry: Entry, ctx: PrintContext, native) -> int | None:
    try:
        result = native(entry.path, ctx.printer, ctx.copies)
    except _PrintViaPdf as via:
        try:
            return _print_pdf(via.pdf, entry.display, ctx)
        finally:
            shutil.rmtree(via.pdf.parent, ignore_errors=True)
    if result is not None or _native_available(ctx, entry.kind):
        return result
    # Sem Microsoft Office: converte com o LibreOffice e imprime o PDF.
    outdir = Path(tempfile.mkdtemp(prefix="printer_conv_"))
    try:
        converted = ctx.office.convert_to_pdf(entry.path, outdir)
        return _print_pdf(converted, entry.display, ctx)
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


def _native_available(ctx: PrintContext, kind: FileKind) -> bool:
    app = ctx.office.word() if kind is FileKind.WORD else ctx.office.excel()
    return app is not None


def _os_message(exc: OSError) -> str:
    winerror = getattr(exc, "winerror", None)
    if winerror == 1801 or "1801" in str(exc):
        return "Impressora não encontrada (verifique o nome ou a conexão de rede)"
    if winerror == 5:
        return "Acesso negado à impressora"
    return exc.strerror or str(exc)
