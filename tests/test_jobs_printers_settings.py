from __future__ import annotations

import queue
import sys
import threading
import types
from pathlib import Path

import pytest

from printer_app import printers
from printer_app.files import Entry, FileKind
from printer_app.jobs import JobEvent, PrintJob, Status
from printer_app.printing import Cancelled, PrintError
from printer_app.settings import Settings, load_settings, save_settings


class NullSession:
    closed = False

    def __enter__(self):
        return self

    def __exit__(self, *a):
        NullSession.closed = True


def _entries(*specs):
    out = []
    for n, (name, kind, error) in enumerate(specs):
        out.append((n, Entry(Path(name), name, "pasta", kind, error)))
    return out


def _drain(q: queue.Queue) -> list[JobEvent]:
    events = []
    while True:
        event = q.get(timeout=5)
        events.append(event)
        if event.index is None:
            return events


def test_job_isolates_errors_and_reports_summary():
    def print_fn(entry, ctx):
        assert ctx.printer == "BRA3" and ctx.copies == 2
        if entry.display == "ruim.pdf":
            raise PrintError("PDF corrompido")
        if entry.display == "bug.pdf":
            raise RuntimeError("boom")
        return 3

    events: queue.Queue = queue.Queue()
    job = PrintJob(
        _entries(
            ("a.pdf", FileKind.PDF, None),
            ("ruim.pdf", FileKind.PDF, None),
            ("x.exe", FileKind.UNSUPPORTED, None),
            ("bug.pdf", FileKind.PDF, None),
            ("b.png", FileKind.IMAGE, None),
        ),
        "BRA3",
        2,
        events,
        print_fn=print_fn,
        session_factory=NullSession,
    )
    job.start()
    got = [(e.index, e.status, e.detail, e.pages) for e in _drain(events)]
    assert got == [
        (0, Status.PRINTING, "", None),
        (0, Status.DONE, "", 3),
        (1, Status.PRINTING, "", None),
        (1, Status.ERROR, "PDF corrompido", None),
        (2, Status.SKIPPED, "Tipo não suportado", None),
        (3, Status.PRINTING, "", None),
        (3, Status.ERROR, "Erro inesperado: boom", None),
        (4, Status.PRINTING, "", None),
        (4, Status.DONE, "", 3),
        (None, None, "", None),
    ]
    s = job.summary
    assert (s.done, s.errors, s.skipped, s.cancelled, s.pages) == (2, 2, 1, 0, 6)
    assert NullSession.closed


def test_job_cancel_stops_remaining_items():
    started = threading.Event()
    release = threading.Event()

    def print_fn(entry, ctx):
        started.set()
        release.wait(5)
        if ctx.cancel.is_set():
            raise Cancelled
        return 1

    events: queue.Queue = queue.Queue()
    job = PrintJob(
        _entries(*[(f"{i}.pdf", FileKind.PDF, None) for i in range(4)]),
        "P",
        1,
        events,
        print_fn=print_fn,
        session_factory=NullSession,
    )
    job.start()
    assert started.wait(5)
    job.cancel()
    release.set()
    statuses = [e.status for e in _drain(events) if e.index is not None]
    assert statuses == [Status.PRINTING] + [Status.CANCELLED] * 4
    assert job.summary.cancelled == 4


def test_job_always_emits_finish_even_if_session_fails():
    class Broken:
        def __enter__(self):
            raise OSError("COM indisponível")

        def __exit__(self, *a):
            pass

    events: queue.Queue = queue.Queue()
    PrintJob(
        _entries(("a.pdf", FileKind.PDF, None), ("b.pdf", FileKind.PDF, None)),
        "P",
        1,
        events,
        print_fn=lambda e, c: 1,
        session_factory=Broken,
    ).start()
    got = [(e.index, e.status, e.detail) for e in _drain(events)]
    assert got == [
        (0, Status.ERROR, "Falha no lote: COM indisponível"),
        (1, Status.ERROR, "Falha no lote: COM indisponível"),
        (None, None, ""),
    ]


# --- impressoras -------------------------------------------------------------------


def test_describe_problem():
    assert printers.describe_problem(0) is None
    assert printers.describe_problem(0x80) == "offline"
    assert printers.describe_problem(0x10 | 0x400000) == "sem papel, com a porta aberta"
    assert printers.describe_problem(0, attributes=0x400) == "offline"
    assert printers.describe_problem(0x400) is None  # "imprimindo" não é problema


@pytest.fixture
def win32print(monkeypatch):
    mod = types.ModuleType("win32print")
    calls = []
    mod.GetDefaultPrinter = lambda: "B"
    mod.EnumPrinters = lambda flags, name, level: [
        {"pPrinterName": "c", "Status": 0, "Attributes": 0},
        {"pPrinterName": "B", "Status": 0, "Attributes": 0},
        {"pPrinterName": "a", "Status": 0x80, "Attributes": 0},
    ]

    def open_printer(name, defaults):
        calls.append(("OpenPrinter", name, defaults))
        return "h"

    mod.OpenPrinter = open_printer
    mod.GetPrinter = lambda h, level: {"Status": 0x10, "Attributes": 0}
    mod.EnumJobs = lambda h, first, count, level: [1, 2]
    mod.ClosePrinter = lambda h: calls.append(("ClosePrinter", h))
    monkeypatch.setitem(sys.modules, "win32print", mod)
    return calls


def test_list_printers_default_first(win32print):
    result = printers.list_printers()
    assert [p.name for p in result] == ["B", "a", "c"]
    assert result[0].is_default and result[1].problem == "offline"


def test_printer_problem_needs_only_use_access(win32print):
    assert printers.printer_problem("X") == "sem papel"
    assert printers.pending_jobs("X") == 2
    # v1 pedia PRINTER_ALL_ACCESS (exige admin); agora só PRINTER_ACCESS_USE.
    assert ("OpenPrinter", "X", {"DesiredAccess": 0x8}) in win32print
    assert win32print.count(("ClosePrinter", "h")) == 2


def test_open_preferences_quotes_name(monkeypatch):
    seen = []
    monkeypatch.setattr(printers.subprocess, "Popen", lambda cmd: seen.append(cmd))
    printers.open_preferences('Recepção "HP" 2º andar')
    assert seen == ['rundll32.exe printui.dll,PrintUIEntry /e /n "Recepção HP 2º andar"']


# --- configurações -----------------------------------------------------------------


def test_settings_roundtrip_and_validation(tmp_path):
    path = tmp_path / "s.json"
    assert load_settings(path) == Settings()
    save_settings(Settings(printer="BRA3", copies=500, appearance="dark"), path)
    assert load_settings(path) == Settings(printer="BRA3", copies=99, appearance="dark")
    path.write_text('{"copies": "abc", "appearance": "neon", "extra": 1}', encoding="utf-8")
    assert load_settings(path) == Settings()
    path.write_text("{corrompido", encoding="utf-8")
    assert load_settings(path) == Settings()
