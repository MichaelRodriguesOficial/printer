"""Testes da janela com o backend de demonstração (rodam no Windows ou em Xvfb no Linux)."""

from __future__ import annotations

import contextlib
import json
import os
import time
import tkinter
import zipfile

import pytest
from PIL import Image

pytestmark = pytest.mark.skipif(
    os.name != "nt" and not os.environ.get("DISPLAY"), reason="sem display gráfico"
)


@pytest.fixture
def make_app(tmp_path, monkeypatch):
    from printer_app.backend import DemoBackend
    from printer_app.ui import app as app_module

    apps = []
    asked = []
    monkeypatch.setattr(
        app_module.messagebox, "askyesno", lambda *a, **k: asked.append(a[1]) or False
    )

    def factory(backend=None, settings=None):
        while apps:  # um Tk por vez, como no app real
            _destroy(apps.pop())
        path = tmp_path / "settings.json"
        if settings is not None:
            path.write_text(json.dumps(settings), encoding="utf-8")
        for attempt in range(3):
            try:
                app = app_module.App(backend or DemoBackend(delay=0), settings_file=path)
                break
            except tkinter.TclError as exc:
                # Falha intermitente conhecida do Tcl/Tk no Windows ao criar vários Tk()
                # seguidos no mesmo processo ("Can't find a usable init.tcl/tk.tcl"). O app real
                # cria um único Tk, então só os testes precisam tentar de novo.
                if "Can't find a usable" not in str(exc) or attempt == 2:
                    raise
                time.sleep(0.5)
        app.asked = asked
        apps.append(app)
        pump(app, lambda: str(app.refresh_button.cget("state")) == "normal")  # lista carregada
        return app

    yield factory
    for app in apps:
        _destroy(app)


def _destroy(app):
    with contextlib.suppress(Exception):
        app.destroy()


def pump(app, condition, timeout=10.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        app.update()
        if condition():
            return
        time.sleep(0.02)
    raise AssertionError("condição não atingida a tempo")


@pytest.fixture
def files(tmp_path, sample_pdf):
    folder = tmp_path / "entrada"
    folder.mkdir()
    sample_pdf("entrada/Pedido 2.pdf", pages=2)
    sample_pdf("entrada/Pedido 10.pdf", pages=1)
    sample_pdf("entrada/Contrato com senha.pdf", pages=1)
    Image.new("RGB", (40, 40)).save(folder / "foto.png")
    (folder / "programa.exe").write_bytes(b"x")
    with zipfile.ZipFile(folder / "lote.zip", "w") as zf:
        zf.writestr("dentro.pdf", (folder / "Pedido 10.pdf").read_bytes())
    return folder


def statuses(app):
    return {app.items[i].entry.display: app.items[i].status for i in app.tree.get_children()}


def test_default_printer_preselected_and_saved_printer_wins(make_app):
    app = make_app()
    assert app.printer_var.get() == "BRA3-ATENDIMENTO"
    assert "padrão do Windows" in app.printer_status.cget("text")
    app2 = make_app(settings={"printer": "BRA3-FINANCEIRO"})
    assert app2.printer_var.get() == "BRA3-FINANCEIRO"


def test_full_flow_add_print_and_summary(make_app, files):
    from printer_app.jobs import Status

    app = make_app()
    assert str(app.print_button.cget("state")) == "disabled"  # lista vazia
    assert app.empty.winfo_ismapped()
    app.add_paths([str(files)])
    pump(app, lambda: app.items and not app._preparing)
    assert [app.items[i].entry.display for i in app.tree.get_children()] == [
        "Contrato com senha.pdf",
        "foto.png",
        "dentro.pdf",
        "Pedido 2.pdf",
        "Pedido 10.pdf",
        "programa.exe",
    ]
    assert statuses(app)["programa.exe"] is Status.SKIPPED
    assert "5 arquivo(s)" in app.count_label.cget("text")
    assert not app.empty.winfo_ismapped()

    app.copies_var.set("2")
    assert app.start_print(confirm=False)
    assert app.print_button.cget("text") == "Cancelar"
    assert str(app.add_files_button.cget("state")) == "disabled"
    pump(app, lambda: app.job is None)

    st = statuses(app)
    assert st["Contrato com senha.pdf"] is Status.ERROR
    assert all(st[n] is Status.DONE for n in ("dentro.pdf", "foto.png", "Pedido 2.pdf"))
    assert app.items[app.tree.get_children()[3]].pages == 4  # Pedido 2: 2 págs x 2 cópias
    assert app.status_label.cget("text").startswith("Concluído: 4 impresso(s)")
    assert "1 com erro" in app.status_label.cget("text")
    assert app.print_button.cget("text") == "Imprimir"
    assert {(p, c) for p, _, c in app.backend.printed} == {("BRA3-ATENDIMENTO", 2)}

    # Reimprimir só reenvia o que falhou.
    app.backend.printed.clear()
    assert app.start_print(confirm=False)
    pump(app, lambda: app.job is None)
    assert app.backend.printed == []  # o único pendente era o PDF com senha
    assert statuses(app)["Contrato com senha.pdf"] is Status.ERROR


def test_cancel_marks_remaining_items(make_app, files):
    from printer_app.backend import DemoBackend
    from printer_app.jobs import Status

    app = make_app(DemoBackend(delay=0.3))
    app.add_paths([str(files)])
    pump(app, lambda: app.items and not app._preparing)
    app.start_print(confirm=False)
    pump(app, lambda: Status.PRINTING in statuses(app).values())
    app.print_or_cancel()
    pump(app, lambda: app.job is None)
    values = list(statuses(app).values())
    assert Status.CANCELLED in values
    assert "cancelado" in app.status_label.cget("text")


def test_offline_printer_asks_before_printing(make_app, files):
    app = make_app()
    app.printer_var.set("Recepção - HP LaserJet")
    app._on_printer_selected("Recepção - HP LaserJet")
    assert "Offline" in app.printer_status.cget("text")
    app.add_paths([str(files / "foto.png")])
    pump(app, lambda: app.items and not app._preparing)
    assert app.start_print(confirm=True) is False  # usuário respondeu "não"
    assert "offline" in app.asked[-1]
    assert app.job is None


def test_remove_clear_and_copies_bounds(make_app, files):
    app = make_app()
    app.add_paths([str(files / "foto.png"), str(files / "lote.zip")])
    pump(app, lambda: len(app.items) == 2 and not app._preparing)
    app.add_paths([str(files / "foto.png")])  # repetido: não entra de novo
    pump(app, lambda: not app._preparing)
    assert len(app.items) == 2
    assert "já estava" in app.status_label.cget("text")
    first = app.tree.get_children()[0]
    app.tree.selection_set(first)
    app.remove_selected()
    assert len(app.items) == 1
    workdir_files = list(app.workdir.rglob("*"))
    assert workdir_files  # zip extraído na pasta temporária da sessão
    app.clear_list()
    assert not app.items and not list(app.workdir.rglob("*"))

    app.copies_var.set("0")
    assert app.copies == 1
    app.copies_var.set("500")
    app._normalize_copies()
    assert app.copies_var.get() == "99"
    app.copies_var.set("98")
    app._step_copies(5)
    assert app.copies_var.get() == "99"


def test_settings_saved_on_close(make_app, tmp_path):
    app = make_app()
    app._on_appearance("Escuro")
    app.fast_pdf_switch.deselect()
    app._on_pdf_mode()
    app.copies_var.set("3")
    app.printer_var.set("BRA3-FINANCEIRO")
    app._on_printer_selected("BRA3-FINANCEIRO")
    workdir = app.workdir
    app.on_close()
    saved = json.loads((tmp_path / "settings.json").read_text(encoding="utf-8"))
    assert saved == {
        "printer": "BRA3-FINANCEIRO",
        "copies": 3,
        "appearance": "dark",
        "pdf_mode": "raster",
    }
    assert not workdir.exists()


def test_no_printers_installed(make_app):
    from printer_app.backend import DemoBackend

    class NoPrinters(DemoBackend):
        def list_printers(self):
            return []

    app = make_app(NoPrinters(delay=0))
    assert "Nenhuma" in app.printer_status.cget("text")
    assert app.selected_printer is None
    assert str(app.print_button.cget("state")) == "disabled"
