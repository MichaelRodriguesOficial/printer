"""Consulta às impressoras do Windows (lista, padrão, status e preferências)."""

from __future__ import annotations

import logging
import subprocess
from dataclasses import dataclass

log = logging.getLogger(__name__)

# Flags de PRINTER_INFO_2.Status (winspool.h) que impedem a impressão.
_STATUS_PROBLEMS: dict[int, str] = {
    0x00000001: "pausada",
    0x00000002: "com erro",
    0x00000004: "sendo removida",
    0x00000008: "com papel atolado",
    0x00000010: "sem papel",
    0x00000040: "com problema no papel",
    0x00000080: "offline",
    0x00000800: "com bandeja de saída cheia",
    0x00001000: "indisponível",
    0x00040000: "sem toner",
    0x00100000: "aguardando intervenção do usuário",
    0x00200000: "sem memória",
    0x00400000: "com a porta aberta",
}
_ATTRIBUTE_WORK_OFFLINE = 0x00000400
_PRINTER_ACCESS_USE = 0x00000008
_ENUM_LOCAL_AND_CONNECTIONS = 0x00000002 | 0x00000004


@dataclass(frozen=True)
class PrinterInfo:
    name: str
    is_default: bool = False
    problem: str | None = None


def describe_problem(status: int, attributes: int = 0) -> str | None:
    """Traduz os flags de status em uma frase curta, ou ``None`` se estiver tudo certo."""
    problems = [text for flag, text in _STATUS_PROBLEMS.items() if status & flag]
    if attributes & _ATTRIBUTE_WORK_OFFLINE and "offline" not in problems:
        problems.insert(0, "offline")
    return ", ".join(problems) if problems else None


def default_printer() -> str | None:
    import win32print

    try:
        return win32print.GetDefaultPrinter() or None
    except Exception:  # noqa: BLE001 - sem impressora padrão configurada
        return None


def list_printers() -> list[PrinterInfo]:
    import win32print

    default = default_printer()
    printers: dict[str, PrinterInfo] = {}
    for info in win32print.EnumPrinters(_ENUM_LOCAL_AND_CONNECTIONS, None, 2):
        name = info["pPrinterName"]
        printers[name] = PrinterInfo(
            name=name,
            is_default=name == default,
            problem=describe_problem(info.get("Status", 0), info.get("Attributes", 0)),
        )
    return sorted(printers.values(), key=lambda p: (not p.is_default, p.name.casefold()))


def printer_problem(name: str) -> str | None:
    """Status atualizado de uma impressora (só exige permissão de uso, não de admin)."""
    import win32print

    handle = win32print.OpenPrinter(name, {"DesiredAccess": _PRINTER_ACCESS_USE})
    try:
        info = win32print.GetPrinter(handle, 2)
    finally:
        win32print.ClosePrinter(handle)
    return describe_problem(info.get("Status", 0), info.get("Attributes", 0))


def pending_jobs(name: str) -> int:
    import win32print

    handle = win32print.OpenPrinter(name, {"DesiredAccess": _PRINTER_ACCESS_USE})
    try:
        return len(win32print.EnumJobs(handle, 0, 999, 1))
    finally:
        win32print.ClosePrinter(handle)


def open_preferences(name: str) -> None:
    """Abre a janela "Preferências de impressão" da impressora, sem travar o app."""
    escaped = name.replace('"', "")
    subprocess.Popen(f'rundll32.exe printui.dll,PrintUIEntry /e /n "{escaped}"')
