"""Caminhos de recursos e de dados do usuário (funciona no código-fonte e no .exe)."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from . import APP_ID

_PACKAGE_ROOT = Path(__file__).resolve().parent.parent


def resource_path(relative: str) -> Path:
    """Caminho absoluto de um recurso empacotado (pasta ``assets``).

    No executável do PyInstaller os dados ficam em ``sys._MEIPASS``; no código-fonte,
    na raiz do projeto. Assim o app não depende da pasta de onde foi aberto.
    """
    base = Path(getattr(sys, "_MEIPASS", _PACKAGE_ROOT))
    return base / relative


def user_data_dir() -> Path:
    """Pasta de configuração do usuário (``%APPDATA%\\PrinterCAV`` no Windows)."""
    if os.name == "nt":
        root = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    path = root / APP_ID
    path.mkdir(parents=True, exist_ok=True)
    return path


def log_dir() -> Path:
    """Pasta de logs (``%LOCALAPPDATA%\\PrinterCAV\\logs`` no Windows)."""
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
        path = root / APP_ID / "logs"
    else:
        path = user_data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path
