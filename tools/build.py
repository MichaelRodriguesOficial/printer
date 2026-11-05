"""Gera o executável ``dist/PrinterCAV.exe`` com o PyInstaller.

Uso (no Windows, dentro da pasta do projeto):
    pip install -r requirements-dev.txt
    python tools/build.py
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from printer_app import APP_ID, APP_NAME, __version__  # noqa: E402

VERSION_TEMPLATE = """VSVersionInfo(
  ffi=FixedFileInfo(filevers={tup}, prodvers={tup}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041604B0', [
      StringStruct('CompanyName', 'CAV SUL - Centro de Apoio a Vendas'),
      StringStruct('FileDescription', '{name}'),
      StringStruct('FileVersion', '{ver}'),
      StringStruct('InternalName', '{app_id}'),
      StringStruct('OriginalFilename', '{app_id}.exe'),
      StringStruct('ProductName', '{name}'),
      StringStruct('ProductVersion', '{ver}'),
      StringStruct('LegalCopyright', 'Michael Rodrigues - Licença MIT')])]),
    VarFileInfo([VarStruct('Translation', [0x0416, 1200])])
  ]
)
"""


def main() -> None:
    import PyInstaller.__main__

    parts = [int(p) for p in __version__.split(".")][:3]
    version_tuple = tuple(parts + [0] * (4 - len(parts)))
    with tempfile.TemporaryDirectory() as tmp:
        version_file = Path(tmp) / "version_info.txt"
        version_file.write_text(
            VERSION_TEMPLATE.format(
                tup=version_tuple, ver=__version__, name=APP_NAME, app_id=APP_ID
            ),
            encoding="utf-8",
        )
        PyInstaller.__main__.run(
            [
                str(ROOT / "main.py"),
                "--name", APP_ID,
                "--onefile",
                "--windowed",
                "--noconfirm",
                "--clean",
                "--icon", str(ROOT / "assets" / "favicon.ico"),
                "--add-data", f"{ROOT / 'assets'}{os.pathsep}assets",
                "--version-file", str(version_file),
                "--collect-data", "customtkinter",
                "--collect-all", "tkinterdnd2",
                "--collect-all", "pypdfium2",
                "--collect-all", "pypdfium2_raw",
                "--collect-submodules", "patoolib",
                "--distpath", str(ROOT / "dist"),
                "--workpath", str(ROOT / "build"),
                "--specpath", str(ROOT / "build"),
            ]
        )  # fmt: skip
    print(f"\nExecutável gerado em {ROOT / 'dist' / (APP_ID + '.exe')}")


if __name__ == "__main__":
    main()
