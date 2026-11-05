from __future__ import annotations

import io
import os
import tarfile
import zipfile
from pathlib import Path

import pytest

from printer_app import archives
from printer_app.archives import ArchiveError, extract_archive
from printer_app.files import Collector, FileKind, is_ignored, kind_of, natural_key


@pytest.mark.parametrize(
    ("name", "kind"),
    [
        ("a.PDF", FileKind.PDF),
        ("foto.JPEG", FileKind.IMAGE),
        ("scan.tif", FileKind.IMAGE),
        ("carta.docx", FileKind.WORD),
        ("antigo.doc", FileKind.WORD),
        ("macro.docm", FileKind.WORD),
        ("tabela.csv", FileKind.EXCEL),
        ("contas.xls", FileKind.EXCEL),
        ("pacote.zip", FileKind.ARCHIVE),
        ("pacote.tar.gz", FileKind.ARCHIVE),
        ("pacote.7z", FileKind.ARCHIVE),
        ("programa.exe", FileKind.UNSUPPORTED),
        ("sem_extensao", FileKind.UNSUPPORTED),
        # Na v1, ".pd" e arquivos sem extensão caíam no fluxo de PDF ("x" in ".pdf").
        ("arquivo.pd", FileKind.UNSUPPORTED),
        ("arquivo.df", FileKind.UNSUPPORTED),
    ],
)
def test_kind_of(name, kind):
    assert kind_of(name) is kind


def test_natural_sort_and_ignored():
    names = ["Boleto 10.pdf", "boleto 2.pdf", "Boleto 1.pdf"]
    assert sorted(names, key=natural_key) == ["Boleto 1.pdf", "boleto 2.pdf", "Boleto 10.pdf"]
    assert is_ignored("~$Relatorio.docx")
    assert is_ignored("Thumbs.db")
    assert is_ignored("__MACOSX")
    assert not is_ignored("Relatorio.docx")


def _zip(path: Path, files: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    return path


def test_collect_folder_recursive_sorted(tmp_path):
    root = tmp_path / "Pedidos"
    (root / "sub").mkdir(parents=True)
    for name in ("p10.pdf", "p2.pdf", "notas.txt", "app.exe", "~$lock.docx", "Thumbs.db"):
        (root / name).write_bytes(b"x")
    (root / "sub" / "foto.png").write_bytes(b"x")
    entries = Collector(tmp_path / "work").collect([root])
    assert [(e.display, e.origin) for e in entries] == [
        ("app.exe", "Pedidos"),
        ("notas.txt", "Pedidos"),
        ("p2.pdf", "Pedidos"),
        ("p10.pdf", "Pedidos"),
        ("foto.png", "Pedidos/sub"),
    ]
    assert [e.printable for e in entries] == [False, True, True, True, True]


def test_collect_two_zips_with_same_names_do_not_collide(tmp_path):
    z1 = _zip(tmp_path / "a.zip", {"doc.pdf": b"AAA"})
    z2 = _zip(tmp_path / "b.zip", {"doc.pdf": b"BBB"})
    entries = Collector(tmp_path / "work").collect([z1, z2])
    assert [e.origin for e in entries] == [f"{tmp_path.name} › a.zip", f"{tmp_path.name} › b.zip"]
    assert [e.path.read_bytes() for e in entries] == [b"AAA", b"BBB"]


def test_collect_zip_with_subfolders_and_nested_zip(tmp_path):
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w") as zf:
        zf.writestr("dentro.pdf", b"x")
    outer = _zip(
        tmp_path / "lote.zip",
        {"pasta/1.pdf": b"x", "2.png": b"x", "interno.zip": inner.getvalue(), "pasta/": b""},
    )
    entries = Collector(tmp_path / "work").collect([outer])
    names = {e.display: e.origin for e in entries}
    assert set(names) == {"1.pdf", "2.png", "dentro.pdf"}
    assert names["dentro.pdf"].endswith("lote.zip › interno.zip")
    assert names["1.pdf"].endswith("lote.zip/pasta")


def test_nested_depth_limit(tmp_path):
    data = b"x"
    name = "final.pdf"
    for level in range(5):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr(name, data)
        data, name = buf.getvalue(), f"n{level}.zip"
    path = tmp_path / "top.zip"
    path.write_bytes(data)
    entries = Collector(tmp_path / "work", max_depth=3).collect([path])
    assert len(entries) == 1
    assert entries[0].error == "Compactado aninhado demais"


def test_missing_and_empty(tmp_path):
    empty = _zip(tmp_path / "vazio.zip", {})
    entries = Collector(tmp_path / "work").collect([tmp_path / "nao_existe.pdf", empty])
    assert entries[0].error == "Arquivo não encontrado"
    assert entries[1].error == "Compactado vazio"
    assert not any(e.printable for e in entries)


def test_zip_windows_explorer_accented_names(tmp_path):
    """Zip do Explorer (cp850, sem flag UTF-8) deve manter os acentos."""
    real = "Relatório Março.pdf"
    placeholder = "Relatorio Marco.pdf"  # mesmo tamanho em bytes que o nome em cp850
    assert len(real.encode("cp850")) == len(placeholder)
    path = _zip(tmp_path / "explorer.zip", {placeholder: b"x"})
    path.write_bytes(path.read_bytes().replace(placeholder.encode(), real.encode("cp850")))
    dest = tmp_path / "out"
    extract_archive(path, dest)
    assert [p.name for p in dest.iterdir()] == [real]


def test_zip_member_name_decoding():
    info = zipfile.ZipInfo("x")
    info.flag_bits = 0
    info.filename = "Relatório.pdf".encode("cp850").decode("cp437")
    assert archives._zip_member_name(info) == "Relatório.pdf"
    info.filename = "Relatório.pdf".encode().decode("cp437")
    assert archives._zip_member_name(info) == "Relatório.pdf"


def test_zip_path_traversal_is_blocked(tmp_path):
    path = _zip(
        tmp_path / "mal.zip",
        {"../fora.pdf": b"x", "C:/Windows/evil.pdf": b"x", "/abs.pdf": b"x", "ok.pdf": b"x"},
    )
    dest = tmp_path / "out"
    extract_archive(path, dest)
    assert not (tmp_path / "fora.pdf").exists()
    files = sorted(p.relative_to(dest).as_posix() for p in dest.rglob("*") if p.is_file())
    assert files == ["abs.pdf", "ok.pdf"]


def test_zip_encrypted_and_corrupted(tmp_path):
    path = _zip(tmp_path / "senha.zip", {"a.pdf": b"x"})
    raw = bytearray(path.read_bytes())
    # liga o bit de criptografia no cabeçalho central
    idx = raw.find(b"PK\x01\x02")
    raw[idx + 8] |= 0x1
    path.write_bytes(bytes(raw))
    with pytest.raises(ArchiveError, match="senha"):
        extract_archive(path, tmp_path / "o1")

    bad = tmp_path / "ruim.zip"
    bad.write_bytes(b"isto nao e um zip")
    with pytest.raises(ArchiveError, match="corrompido"):
        extract_archive(bad, tmp_path / "o2")
    assert not (tmp_path / "o2").exists()


def test_zip_limits(tmp_path, monkeypatch):
    monkeypatch.setattr(archives, "MAX_FILES", 2)
    path = _zip(tmp_path / "muitos.zip", {f"{i}.pdf": b"x" * 8 for i in range(3)})
    with pytest.raises(ArchiveError, match="arquivos demais"):
        extract_archive(path, tmp_path / "out")
    monkeypatch.setattr(archives, "MAX_FILES", 100)
    monkeypatch.setattr(archives, "MAX_TOTAL_BYTES", 10)
    with pytest.raises(ArchiveError, match="grande demais"):
        extract_archive(path, tmp_path / "out")


def test_tar_extraction_filters_unsafe_members(tmp_path):
    path = tmp_path / "pacote.tar.gz"
    with tarfile.open(path, "w:gz") as tf:
        for name in ("ok.pdf", "../fora.pdf"):
            info = tarfile.TarInfo(name)
            info.size = 1
            tf.addfile(info, io.BytesIO(b"x"))
    dest = tmp_path / "out"
    with pytest.raises(ArchiveError):
        extract_archive(path, dest)  # filtro "data" recusa o membro malicioso
    assert not (tmp_path / "fora.pdf").exists()

    safe = tmp_path / "seguro.tar"
    with tarfile.open(safe, "w") as tf:
        info = tarfile.TarInfo("pasta/ok.pdf")
        info.size = 1
        tf.addfile(info, io.BytesIO(b"x"))
    extract_archive(safe, tmp_path / "out2")
    assert (tmp_path / "out2" / "pasta" / "ok.pdf").exists()


def test_rar_without_7zip_gives_clear_message(tmp_path, monkeypatch):
    monkeypatch.setenv("PATH", str(tmp_path))  # nenhum descompactador disponível
    path = tmp_path / "x.rar"
    path.write_bytes(b"Rar!\x1a\x07\x00" + b"\x00" * 32)
    with pytest.raises(ArchiveError, match="7-Zip"):
        extract_archive(path, tmp_path / "out")


@pytest.mark.skipif(os.name == "nt", reason="symlink exige privilégio no Windows")
def test_zip_symlink_member_skipped(tmp_path):
    path = tmp_path / "link.zip"
    with zipfile.ZipFile(path, "w") as zf:
        info = zipfile.ZipInfo("link.pdf")
        info.external_attr = 0o120777 << 16
        zf.writestr(info, "/etc/passwd")
        zf.writestr("real.pdf", b"x")
    dest = tmp_path / "out"
    extract_archive(path, dest)
    assert sorted(p.name for p in dest.iterdir()) == ["real.pdf"]
