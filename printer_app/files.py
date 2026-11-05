"""Classificação de arquivos e montagem da lista de impressão (arquivos, pastas e compactados)."""

from __future__ import annotations

import itertools
import logging
import os
import re
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from .archives import ArchiveError, extract_archive

log = logging.getLogger(__name__)


class FileKind(Enum):
    PDF = "PDF"
    IMAGE = "Imagem"
    WORD = "Word"
    EXCEL = "Excel"
    ARCHIVE = "Compactado"
    UNSUPPORTED = "Não suportado"

    @property
    def printable(self) -> bool:
        return self in (FileKind.PDF, FileKind.IMAGE, FileKind.WORD, FileKind.EXCEL)


_KIND_BY_EXT: dict[str, FileKind] = {
    ".pdf": FileKind.PDF,
    **dict.fromkeys(
        (".png", ".jpg", ".jpeg", ".jfif", ".bmp", ".gif", ".tif", ".tiff", ".webp"),
        FileKind.IMAGE,
    ),
    **dict.fromkeys((".doc", ".docx", ".docm", ".rtf", ".odt", ".txt"), FileKind.WORD),
    **dict.fromkeys((".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods"), FileKind.EXCEL),
    **dict.fromkeys(
        (".zip", ".rar", ".7z", ".tar", ".tgz", ".gz", ".bz2", ".xz"), FileKind.ARCHIVE
    ),
}

# Arquivos de sistema/temporários que nunca devem ir para a impressora.
_IGNORED_NAMES = {"thumbs.db", "desktop.ini", ".ds_store", "__macosx"}

MAX_ARCHIVE_DEPTH = 3


def kind_of(path: str | os.PathLike[str]) -> FileKind:
    return _KIND_BY_EXT.get(Path(path).suffix.lower(), FileKind.UNSUPPORTED)


def is_ignored(name: str) -> bool:
    lower = name.lower()
    return lower in _IGNORED_NAMES or lower.startswith(("~$", "."))


def _folder_label(folder: Path) -> str:
    return folder.name or str(folder)


def natural_key(text: str) -> list[object]:
    """Ordenação "humana": ``arquivo2`` vem antes de ``arquivo10``."""
    return [int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", text)]


@dataclass
class Entry:
    """Um item da lista de impressão."""

    path: Path
    display: str
    origin: str
    kind: FileKind
    error: str | None = None

    @property
    def printable(self) -> bool:
        return self.kind.printable and self.error is None


class Collector:
    """Expande arquivos, pastas e compactados em uma lista ordenada de :class:`Entry`.

    Cada compactado é extraído em sua própria subpasta de ``workdir`` — dois .zip com
    arquivos de mesmo nome não se sobrescrevem mais.
    """

    def __init__(self, workdir: Path, max_depth: int = MAX_ARCHIVE_DEPTH) -> None:
        self.workdir = workdir
        self.max_depth = max_depth
        self._counter = itertools.count(1)

    def collect(self, paths: Iterable[str | os.PathLike[str]]) -> list[Entry]:
        entries: list[Entry] = []
        for raw in paths:
            path = Path(raw)
            if path.is_dir():
                entries.extend(self._walk_dir(path, origin=path.name, depth=0))
            elif path.exists():
                entries.extend(self._file(path, origin=_folder_label(path.parent), depth=0))
            else:
                entries.append(
                    Entry(
                        path,
                        path.name,
                        _folder_label(path.parent),
                        kind_of(path),
                        "Arquivo não encontrado",
                    )
                )
        return entries

    def _walk_dir(self, root: Path, origin: str, depth: int) -> Iterator[Entry]:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = sorted((d for d in dirnames if not is_ignored(d)), key=natural_key)
            rel_dir = Path(dirpath).relative_to(root)
            for name in sorted(filenames, key=natural_key):
                if is_ignored(name):
                    continue
                sub_origin = origin if rel_dir == Path(".") else f"{origin}/{rel_dir.as_posix()}"
                yield from self._file(Path(dirpath) / name, sub_origin, depth)

    def _file(self, path: Path, origin: str, depth: int) -> Iterator[Entry]:
        kind = kind_of(path)
        if kind is not FileKind.ARCHIVE:
            yield Entry(path, path.name, origin, kind)
            return
        if depth >= self.max_depth:
            yield Entry(path, path.name, origin, kind, "Compactado aninhado demais")
            return
        dest = self.workdir / f"{next(self._counter):04d}_{path.stem[:40]}"
        try:
            extract_archive(path, dest)
        except ArchiveError as exc:
            log.warning("Falha ao extrair %s: %s", path, exc)
            yield Entry(path, path.name, origin, kind, str(exc))
            return
        inner = list(self._walk_dir(dest, origin=f"{origin} › {path.name}", depth=depth + 1))
        if not inner:
            yield Entry(path, path.name, origin, kind, "Compactado vazio")
        yield from inner
