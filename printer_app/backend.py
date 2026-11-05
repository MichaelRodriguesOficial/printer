"""Ponte entre a interface e o sistema: Windows de verdade ou modo demonstração."""

from __future__ import annotations

import random
import time
from pathlib import Path
from typing import Protocol

from . import printers
from .files import Entry, FileKind
from .jobs import PrintFn
from .printers import PrinterInfo
from .printing import PrintContext, PrintError, print_entry
from .printing.gdi import Cancelled
from .printing.office import OfficeSession


class Backend(Protocol):
    print_fn: PrintFn

    def list_printers(self) -> list[PrinterInfo]: ...
    def printer_problem(self, name: str) -> str | None: ...
    def open_preferences(self, name: str) -> None: ...
    def session_factory(self) -> OfficeSession: ...


class WindowsBackend:
    print_fn: PrintFn = staticmethod(print_entry)

    def list_printers(self) -> list[PrinterInfo]:
        return printers.list_printers()

    def printer_problem(self, name: str) -> str | None:
        return printers.printer_problem(name)

    def open_preferences(self, name: str) -> None:
        printers.open_preferences(name)

    def session_factory(self) -> OfficeSession:
        return OfficeSession()


class _NullSession(OfficeSession):
    def close(self) -> None:
        pass


class DemoBackend:
    """Simula impressoras e impressão — para desenvolvimento, testes e screenshots."""

    def __init__(self, delay: float = 0.35, seed: int | None = 7) -> None:
        self.delay = delay
        self._rng = random.Random(seed)
        self.printed: list[tuple[str, Path, int]] = []

    def list_printers(self) -> list[PrinterInfo]:
        return [
            PrinterInfo("BRA3-ATENDIMENTO", is_default=True),
            PrinterInfo("BRA3-FINANCEIRO"),
            PrinterInfo("Recepção - HP LaserJet", problem="offline"),
            PrinterInfo("Microsoft Print to PDF"),
        ]

    def printer_problem(self, name: str) -> str | None:
        return next((p.problem for p in self.list_printers() if p.name == name), None)

    def open_preferences(self, name: str) -> None:
        pass

    def session_factory(self) -> OfficeSession:
        return _NullSession()

    def print_fn(self, entry: Entry, ctx: PrintContext) -> int | None:
        if entry.error:
            raise PrintError(entry.error)
        steps = max(1, int(self.delay / 0.05))
        for _ in range(steps):
            if ctx.cancel.is_set():
                raise Cancelled
            time.sleep(self.delay / steps)
        if "senha" in entry.display.lower():
            raise PrintError("PDF protegido por senha")
        pages = self._pages(entry)
        self.printed.append((ctx.printer, entry.path, ctx.copies))
        return None if pages is None else pages * ctx.copies

    def _pages(self, entry: Entry) -> int | None:
        if entry.kind is FileKind.PDF:
            from .printing import pdf

            try:
                return pdf.page_count(entry.path)
            except Exception:  # noqa: BLE001 - arquivos de demonstração podem ser falsos
                return self._rng.randint(1, 12)
        if entry.kind is FileKind.IMAGE:
            return 1
        if entry.kind is FileKind.WORD:
            return self._rng.randint(1, 6)
        return None
