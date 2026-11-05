"""Gera as capturas de tela do README usando o modo demonstração.

Uso: python tools/screenshot.py [pasta_de_saida]
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PIL import Image, ImageDraw, ImageGrab  # noqa: E402

from printer_app.backend import DemoBackend  # noqa: E402
from printer_app.ui.app import App  # noqa: E402


def make_samples(folder: Path) -> list[str]:
    def pdf(name: str, pages: int) -> Path:
        imgs = [Image.new("RGB", (595, 842), "white") for _ in range(pages)]
        path = folder / name
        imgs[0].save(path, "PDF", save_all=True, append_images=imgs[1:])
        return path

    pdf("Pedido 1045 - Consultora Ana.pdf", 2)
    pdf("Pedido 1046 - Consultora Bruna.pdf", 1)
    pdf("Relatório de vendas outubro.pdf", 6)
    pdf("Contrato protegido por senha.pdf", 3)
    img = Image.new("RGB", (800, 600), "#e8eef8")
    ImageDraw.Draw(img).rectangle((80, 80, 720, 520), outline="#2f5da8", width=12)
    img.save(folder / "Comprovante PIX.png")
    (folder / "Planilha de metas.xlsx").write_bytes(b"demo")
    (folder / "Comunicado às consultoras.docx").write_bytes(b"demo")
    (folder / "anotações.exe").write_bytes(b"demo")
    with zipfile.ZipFile(folder / "Boletos novembro.zip", "w") as zf:
        for n in (1, 2, 3):
            boleto = (folder / "Pedido 1046 - Consultora Bruna.pdf").read_bytes()
            zf.writestr(f"Boleto {n:02d}.pdf", boleto)
    return [str(p) for p in sorted(folder.iterdir())]


def grab(app: App, path: Path) -> None:
    app.update()
    x, y = app.winfo_rootx(), app.winfo_rooty()
    w, h = app.winfo_width(), app.winfo_height()
    ImageGrab.grab(
        bbox=(x, y, x + w, y + h), xdisplay=None if sys.platform == "win32" else ""
    ).save(path)
    print("salvo", path)


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "docs" / "screenshots"
    out.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp())
    samples = make_samples(tmp)
    settings = tmp / "settings.json"
    settings.write_text('{"appearance": "light", "copies": 1}', encoding="utf-8")
    backend = DemoBackend(delay=0.25)

    def steps(app: App) -> None:
        app.geometry("1040x640+20+20")

        def s1():
            grab(app, out / "vazio-claro.png")
            app.add_paths(samples)
            app.after(1200, s2)

        def s2():
            app.start_print(confirm=False)
            app.after(1300, s3)

        def s3():
            grab(app, out / "imprimindo-claro.png")
            app.after(4000, s4)

        def s4():
            app._on_appearance("Escuro")
            app.after(800, s5)

        def s5():
            grab(app, out / "concluido-escuro.png")
            app.on_close()

        app.after(1200, s1)

    App(backend, settings_file=settings, on_ready=steps).mainloop()


if __name__ == "__main__":
    main()
