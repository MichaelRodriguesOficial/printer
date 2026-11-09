"""Execução do lote de impressão em segundo plano, com eventos para a interface.

A thread de trabalho nunca toca na interface (Tkinter não é thread-safe): ela publica
eventos em uma ``queue.Queue`` que a janela consome com ``after()``.
"""

from __future__ import annotations

import logging
import queue
import threading
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from enum import Enum

from .files import Entry
from .printing import Cancelled, PrintContext, PrintError, print_entry
from .printing.gdi import PrinterDC
from .printing.office import OfficeSession

log = logging.getLogger(__name__)


class Status(Enum):
    PENDING = "Na fila"
    PRINTING = "Imprimindo…"
    DONE = "Impresso"
    ERROR = "Erro"
    SKIPPED = "Ignorado"
    CANCELLED = "Cancelado"


@dataclass(frozen=True)
class JobEvent:
    """Mudança de status de um item (``index``) ou fim do lote (``index is None``)."""

    index: int | None
    status: Status | None = None
    detail: str = ""
    pages: int | None = None


@dataclass
class Summary:
    done: int = 0
    errors: int = 0
    skipped: int = 0
    cancelled: int = 0
    pages: int = 0
    seconds: float = 0.0


PrintFn = Callable[[Entry, PrintContext], "int | None"]


class PrintJob:
    def __init__(
        self,
        items: Sequence[tuple[int, Entry]],
        printer: str,
        copies: int,
        events: queue.Queue[JobEvent],
        print_fn: PrintFn = print_entry,
        session_factory: Callable[[], OfficeSession] = OfficeSession,
        pdf_mode: str = "vector",
    ) -> None:
        self.items = list(items)
        self.printer = printer
        self.copies = copies
        self.events = events
        self.print_fn = print_fn
        self.session_factory = session_factory
        self.pdf_mode = pdf_mode
        self.cancel_event = threading.Event()
        self.summary = Summary()
        self._thread = threading.Thread(target=self._run, name="print-job", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def cancel(self) -> None:
        self.cancel_event.set()

    def join(self, timeout: float | None = None) -> None:
        self._thread.join(timeout)

    @property
    def running(self) -> bool:
        return self._thread.is_alive()

    def _emit(self, *args, **kwargs) -> None:
        self.events.put(JobEvent(*args, **kwargs))

    def _run(self) -> None:
        log.info(
            "Lote iniciado: %d itens -> %r (%d cópia(s))",
            len(self.items),
            self.printer,
            self.copies,
        )
        remaining = list(self.items)
        started = time.perf_counter()
        printer_dc = PrinterDC(self.printer)  # criado só quando o 1º arquivo GDI precisar
        try:
            with self.session_factory() as session:
                ctx = PrintContext(
                    self.printer, self.copies, self.cancel_event, session, self.pdf_mode,
                    printer_dc,
                )  # fmt: skip
                while remaining:
                    index, entry = remaining[0]
                    if self.cancel_event.is_set():
                        self.summary.cancelled += 1
                        self._emit(index, Status.CANCELLED)
                    else:
                        self._print_one(index, entry, ctx)
                    remaining.pop(0)
        except Exception as exc:
            log.exception("Falha inesperada no lote")
            for index, _entry in remaining:
                self.summary.errors += 1
                self._emit(index, Status.ERROR, f"Falha no lote: {exc}")
        finally:
            try:
                printer_dc.close()
            except Exception:
                log.exception("Falha ao liberar a impressora")
            self.summary.seconds = time.perf_counter() - started
            log.info("Lote finalizado: %s", self.summary)
            self._emit(None)

    def _print_one(self, index: int, entry: Entry, ctx: PrintContext) -> None:
        if not entry.printable:
            self.summary.skipped += 1
            self._emit(index, Status.SKIPPED, entry.error or "Tipo não suportado")
            return
        self._emit(index, Status.PRINTING)
        started = time.perf_counter()
        try:
            pages = self.print_fn(entry, ctx)
        except Cancelled:
            self.summary.cancelled += 1
            self._emit(index, Status.CANCELLED)
        except PrintError as exc:
            log.warning("Erro ao imprimir %s: %s", entry.path, exc)
            self.summary.errors += 1
            self._emit(index, Status.ERROR, str(exc))
        except Exception as exc:
            log.exception("Erro inesperado ao imprimir %s", entry.path)
            self.summary.errors += 1
            self._emit(index, Status.ERROR, f"Erro inesperado: {exc}")
        else:
            log.info(
                "Impresso %s: %s página(s) em %.2f s (PDF: %s)",
                entry.display, pages, time.perf_counter() - started, ctx.pdf_mode,
            )  # fmt: skip
            self.summary.done += 1
            self.summary.pages += pages or 0
            self._emit(index, Status.DONE, pages=pages)
