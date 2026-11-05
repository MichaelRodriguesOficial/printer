"""Atalho para executar o app: ``python main.py`` (também é a entrada do PyInstaller)."""

import sys

from printer_app.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
