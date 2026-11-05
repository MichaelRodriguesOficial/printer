"""Ponto de entrada: ``python -m printer_app`` (ou ``PrinterCAV.exe``).

Opções:
  --demo        impressoras e impressão simuladas (desenvolvimento/apresentação)
  --self-test   verifica recursos, PDFium e a janela e sai (usado no CI após o build)
"""

from __future__ import annotations

import argparse
import io
import logging
import os
import sys
import tempfile
import traceback
from pathlib import Path

from . import APP_NAME, __version__

log = logging.getLogger("printer_app")


def _self_test() -> int:
    """Checagem rápida do executável empacotado: recursos, PDFium, Pillow e Tk."""
    from PIL import Image

    from .paths import resource_path
    from .printing import images, pdf

    for asset in ("assets/logo.png", "assets/favicon.ico"):
        if not resource_path(asset).is_file():
            raise FileNotFoundError(asset)
    with tempfile.TemporaryDirectory() as tmp:
        sample = Path(tmp) / "teste.pdf"
        Image.new("RGB", (200, 300), "white").save(sample, "PDF")
        pages = list(pdf.pdf_pages(sample, 72))
        assert len(pages) == 1 and pages[0].mode == "RGB", "PDFium falhou"
        png = Path(tmp) / "teste.png"
        Image.new("RGBA", (20, 20), (0, 0, 0, 0)).save(png)
        assert next(images.image_pages(png)).getpixel((0, 0)) == (255, 255, 255)

    from .backend import DemoBackend
    from .ui.app import App

    app = App(
        DemoBackend(delay=0),
        settings_file=Path(tempfile.gettempdir()) / "pcav_selftest.json",
        on_ready=lambda a: a.after(200, a.on_close),
    )
    app.mainloop()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="PrinterCAV", description=APP_NAME)
    parser.add_argument("--demo", action="store_true", help="modo demonstração")
    parser.add_argument("--self-test", action="store_true", help="autoteste e sai")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    from .logging_setup import setup_logging

    setup_logging()
    log.info("%s %s iniciando (Python %s)", APP_NAME, __version__, sys.version.split()[0])

    if args.self_test:
        try:
            return _self_test()
        except Exception:  # noqa: BLE001
            _report_fatal()
            return 1

    if os.name != "nt" and not args.demo:
        print("Este aplicativo imprime usando o Windows. Use --demo para testar a interface.")
        return 2

    try:
        from .backend import DemoBackend, WindowsBackend
        from .ui.app import App

        backend = DemoBackend() if args.demo else WindowsBackend()
        App(backend).mainloop()
    except Exception:  # noqa: BLE001
        _report_fatal()
        return 1
    return 0


def _report_fatal() -> None:
    """Registra o erro e o mostra numa caixa de diálogo (o .exe não tem console)."""
    log.exception("Erro fatal")
    text = io.StringIO()
    traceback.print_exc(file=text)
    print(text.getvalue(), file=sys.stderr)
    try:
        from tkinter import Tk, messagebox

        root = Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, f"Erro inesperado:\n\n{text.getvalue()[-1500:]}")
        root.destroy()
    except Exception:  # noqa: BLE001
        pass


if __name__ == "__main__":
    sys.exit(main())
