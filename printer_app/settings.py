"""Preferências persistidas entre execuções (última impressora, cópias, tema)."""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, fields
from pathlib import Path

from .paths import user_data_dir

log = logging.getLogger(__name__)

APPEARANCES = ("system", "light", "dark")
PDF_MODES = ("vector", "raster")
MAX_COPIES = 99


@dataclass
class Settings:
    printer: str | None = None
    copies: int = 1
    appearance: str = "system"
    pdf_mode: str = "vector"

    def normalized(self) -> Settings:
        copies = self.copies if isinstance(self.copies, int) else 1
        return Settings(
            printer=self.printer if isinstance(self.printer, str) and self.printer else None,
            copies=min(max(copies, 1), MAX_COPIES),
            appearance=self.appearance if self.appearance in APPEARANCES else "system",
            pdf_mode=self.pdf_mode if self.pdf_mode in PDF_MODES else "vector",
        )


def settings_path() -> Path:
    return user_data_dir() / "settings.json"


def load_settings(path: Path | None = None) -> Settings:
    path = path or settings_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return Settings()
    except (OSError, ValueError) as exc:
        log.warning("Configurações ilegíveis em %s: %s", path, exc)
        return Settings()
    if not isinstance(data, dict):
        return Settings()
    known = {f.name for f in fields(Settings)}
    return Settings(**{k: v for k, v in data.items() if k in known}).normalized()


def save_settings(settings: Settings, path: Path | None = None) -> None:
    path = path or settings_path()
    try:
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(settings.normalized()), indent=2), encoding="utf-8")
        tmp.replace(path)
    except OSError as exc:
        log.warning("Não foi possível salvar configurações: %s", exc)
