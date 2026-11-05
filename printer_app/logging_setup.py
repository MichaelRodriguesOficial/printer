"""Log em arquivo rotativo — útil para o suporte de TI investigar falhas de impressão."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .paths import log_dir


def setup_logging(level: int = logging.INFO) -> Path | None:
    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s [%(threadName)s] %(name)s: %(message)s")
    try:
        path = log_dir() / "printer.log"
        handler = RotatingFileHandler(path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    except OSError:
        handler, path = logging.StreamHandler(), None
    handler.setFormatter(fmt)
    root.addHandler(handler)
    return path
