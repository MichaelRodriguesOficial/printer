"""Extração segura de compactados.

* ``.zip`` e ``.tar*`` com a biblioteca padrão (proteção contra *path traversal* e
  "zip bomb"; nomes acentuados de zips criados pelo Windows Explorer são corrigidos).
* ``.rar`` e ``.7z`` via ``patool``, que usa o 7-Zip/WinRAR instalado na máquina.
"""

from __future__ import annotations

import os
import shutil
import stat
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

MAX_FILES = 5_000
MAX_TOTAL_BYTES = 2 * 1024**3  # 2 GiB descompactados

_TAR_SUFFIXES = (".tar", ".tgz", ".tar.gz", ".tar.bz2", ".tar.xz", ".gz", ".bz2", ".xz")


class ArchiveError(Exception):
    """Erro de extração com mensagem pronta para o usuário."""


def extract_archive(archive: Path, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    name = archive.name.lower()
    try:
        if name.endswith(".zip"):
            _extract_zip(archive, dest)
        elif name.endswith(_TAR_SUFFIXES):
            _extract_tar(archive, dest)
        elif name.endswith((".rar", ".7z")):
            _extract_with_patool(archive, dest)
        else:
            raise ArchiveError("Formato de compactado não suportado")
    except ArchiveError:
        shutil.rmtree(dest, ignore_errors=True)
        raise
    except (OSError, zipfile.BadZipFile, tarfile.TarError, EOFError) as exc:
        shutil.rmtree(dest, ignore_errors=True)
        raise ArchiveError(f"Compactado corrompido ou ilegível ({exc})") from exc


# --- zip -------------------------------------------------------------------


def _zip_member_name(info: zipfile.ZipInfo) -> str:
    """Nome correto do membro, inclusive para zips do Windows Explorer.

    Sem o flag UTF-8 (bit 11) o Python decodifica como cp437, mas o Explorer em
    português grava em cp850 — "Relatório.pdf" viraria "Relat¾rio.pdf".
    """
    if info.flag_bits & 0x800:
        return info.filename
    raw = info.filename.encode("cp437")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp850")


def _safe_target(dest: Path, member: str) -> Path | None:
    parts = [p for p in PurePosixPath(member.replace("\\", "/")).parts if p not in ("", ".", "/")]
    if not parts or ".." in parts or ":" in parts[0]:
        return None
    target = dest.joinpath(*parts).resolve()
    return target if target.is_relative_to(dest.resolve()) else None


def _extract_zip(archive: Path, dest: Path) -> None:
    with zipfile.ZipFile(archive) as zf:
        infos = [i for i in zf.infolist() if not i.is_dir()]
        if any(i.flag_bits & 0x1 for i in infos):
            raise ArchiveError("Compactado protegido por senha")
        _check_limits(len(infos), sum(i.file_size for i in infos))
        for info in infos:
            mode = info.external_attr >> 16
            if stat.S_ISLNK(mode):
                continue  # links simbólicos são ignorados
            target = _safe_target(dest, _zip_member_name(info))
            if target is None:
                continue  # caminho malicioso (../, C:\...) ignorado
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out, 1024 * 1024)


# --- tar -------------------------------------------------------------------


def _extract_tar(archive: Path, dest: Path) -> None:
    with tarfile.open(archive) as tf:
        members = [m for m in tf.getmembers() if m.isfile()]
        _check_limits(len(members), sum(m.size for m in members))
        tf.extractall(dest, members=members, filter="data")


# --- rar / 7z ----------------------------------------------------------------


def _add_archivers_to_path() -> None:
    """O instalador do 7-Zip/WinRAR não coloca o programa no PATH; o patool precisa dele."""
    if os.name != "nt":
        return
    candidates = []
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        base = os.environ.get(env)
        if base:
            candidates += [Path(base) / "7-Zip", Path(base) / "WinRAR"]
    current = os.environ.get("PATH", "").split(os.pathsep)
    extra = [str(p) for p in candidates if p.is_dir() and str(p) not in current]
    if extra:
        os.environ["PATH"] = os.pathsep.join(current + extra)


def _extract_with_patool(archive: Path, dest: Path) -> None:
    import patoolib
    from patoolib.util import PatoolError

    _add_archivers_to_path()
    try:
        patoolib.extract_archive(str(archive), outdir=str(dest), verbosity=-1, interactive=False)
    except PatoolError as exc:
        msg = str(exc).lower()
        if "could not find an executable" in msg or "not found" in msg:
            raise ArchiveError(
                f"Para abrir {archive.suffix.lower()} instale o 7-Zip (7-zip.org)"
            ) from exc
        raise ArchiveError(f"Não foi possível extrair ({exc})") from exc
    # Validação pós-extração: remove links e qualquer coisa fora do destino.
    root = dest.resolve()
    count = total = 0
    for path in list(dest.rglob("*")):
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            path.unlink(missing_ok=True)
        elif path.is_file():
            count += 1
            total += path.stat().st_size
    _check_limits(count, total)


def _check_limits(count: int, total: int) -> None:
    if count > MAX_FILES:
        raise ArchiveError(f"Compactado com arquivos demais ({count})")
    if total > MAX_TOTAL_BYTES:
        raise ArchiveError("Compactado grande demais para impressão")
