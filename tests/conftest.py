"""Fixtures comuns: módulos falsos do pywin32 para testar a lógica fora do Windows."""

from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))


class FakeDC:
    """Imita o device context do win32ui, registrando as chamadas."""

    def __init__(self, log: list, fail_on_page: int | None = None) -> None:
        self.log = log
        self.fail_on_page = fail_on_page
        self.pages = 0

    def CreatePrinterDC(self, name):
        self.log.append(("CreatePrinterDC", name))

    def GetDeviceCaps(self, index):
        return {8: 2400, 10: 3400, 88: 300}[index]

    def StartDoc(self, name):
        self.log.append(("StartDoc", name))

    def StartPage(self):
        self.pages += 1
        if self.fail_on_page == self.pages:
            raise RuntimeError("falha simulada no driver")
        self.log.append(("StartPage",))

    def EndPage(self):
        self.log.append(("EndPage",))

    def EndDoc(self):
        self.log.append(("EndDoc",))

    def AbortDoc(self):
        self.log.append(("AbortDoc",))

    def DeleteDC(self):
        self.log.append(("DeleteDC",))

    def GetHandleOutput(self):
        return 1234

    def GetSafeHdc(self):
        return 5678


@pytest.fixture
def fake_gdi(monkeypatch):
    """Substitui win32ui e o desenho do Dib; devolve a lista de chamadas e opções."""
    calls: list = []
    state = {"fail_on_page": None, "draws": []}

    module = types.ModuleType("win32ui")
    module.CreateDC = lambda: FakeDC(calls, state["fail_on_page"])
    monkeypatch.setitem(sys.modules, "win32ui", module)

    from PIL import ImageWin

    class FakeDib:
        def __init__(self, image):
            self.image = image

        def draw(self, handle, box):
            state["draws"].append((self.image.size, box))

    monkeypatch.setattr(ImageWin, "Dib", FakeDib)

    import pypdfium2.raw as pdfium_c

    def render_page(hdc, page, x, y, w, h, rotate, flags):
        state["vector"].append((hdc, (x, y, w, h), rotate, flags))
        calls.append(("RenderPage",))

    state["vector"] = []
    state["install_vector"] = lambda: monkeypatch.setattr(
        pdfium_c, "FPDF_RenderPage", render_page, raising=False
    )
    state["calls"] = calls
    return state


@pytest.fixture
def sample_pdf(tmp_path) -> Path:
    def make(name: str = "doc.pdf", pages: int = 3, size=(595, 842)) -> Path:
        imgs = [Image.new("RGB", size, "white") for _ in range(pages)]
        path = tmp_path / name
        imgs[0].save(path, "PDF", save_all=True, append_images=imgs[1:])
        return path

    return make
