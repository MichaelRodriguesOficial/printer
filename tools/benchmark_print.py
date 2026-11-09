"""Compara a impressão de PDF em bitmap (raster) e vetorial numa impressora real ou virtual.

Uso (Windows):
    python tools/benchmark_print.py "Nome da impressora" [--paginas 20] [--saida C:\\saida.pdf]

Atenção: imprime de verdade (duas vezes o documento de teste). Para não gastar papel, use
uma impressora virtual, como a "Microsoft Print to PDF" com porta apontando para arquivo
(é assim que o CI roda). ``--saida`` informa o arquivo gerado por essa porta, para medir
o tamanho do resultado.
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from pdfgen import make_text_pdf  # noqa: E402

from printer_app.printing import gdi, pdf  # noqa: E402


def _wait_queue(printer: str, output: Path | None, timeout: float = 300) -> None:
    from printer_app.printers import pending_jobs

    deadline = time.monotonic() + timeout
    last_size = -1
    while time.monotonic() < deadline:
        if pending_jobs(printer) == 0:
            if output is None:
                return
            size = output.stat().st_size if output.exists() else -1
            if size > 0 and size == last_size:
                return
            last_size = size
        time.sleep(0.25)
    raise TimeoutError("o trabalho não saiu da fila a tempo")


def _has_text(output: Path | None) -> str:
    """Se a saída é PDF, diz se o texto continuou selecionável (sinal de envio vetorial)."""
    if output is None or output.suffix.lower() != ".pdf":
        return "—"
    import pypdfium2 as pdfium

    try:
        doc = pdfium.PdfDocument(str(output))
        try:
            text = doc[0].get_textpage().get_text_range()
        finally:
            doc.close()
    except Exception:  # noqa: BLE001
        return "?"
    return "sim" if "Consultora" in text else "não"


def _send(printer: str, files: list[Path], mode: str, reuse_dc: bool) -> int:
    session = gdi.PrinterDC(printer) if reuse_dc else None
    target = session or printer
    printed = 0
    try:
        for path in files:
            if mode == "vector":
                printed += gdi.print_pdf_vector(target, path, path.name)
            else:
                printed += gdi.print_pages(
                    target, path.name, lambda dpi, p=path: pdf.pdf_pages(p, dpi)
                )
    finally:
        if session is not None:
            session.close()
    return printed


def run(printer: str, pages: int, output: Path | None) -> list[dict]:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        big = [make_text_pdf(tmp / "relatorio.pdf", pages=pages)]
        many = [make_text_pdf(tmp / f"boleto{i:02d}.pdf", pages=1) for i in range(pages)]
        cases = [
            (f"1 PDF de {pages} págs", "v2.0.0: imagem", big, "raster", False),
            (f"1 PDF de {pages} págs", "novo: vetorial", big, "vector", True),
            (
                f"{pages} PDFs de 1 pág",
                "v2.0.0: imagem, 1 conexão por arquivo",
                many,
                "raster",
                False,
            ),
            (f"{pages} PDFs de 1 pág", "imagem + conexão reaproveitada", many, "raster", True),
            (
                f"{pages} PDFs de 1 pág",
                "novo: vetorial + conexão reaproveitada",
                many,
                "vector",
                True,
            ),
        ]
        results = []
        for scenario, name, files, mode, reuse in cases:
            if output is not None and output.exists():
                output.unlink()
            start = time.perf_counter()
            printed = _send(printer, files, mode, reuse)
            sent = time.perf_counter() - start
            _wait_queue(printer, output if len(files) == 1 else None)
            total = time.perf_counter() - start
            single = output is not None and len(files) == 1
            results.append(
                {
                    "cenario": scenario,
                    "modo": name,
                    "paginas": printed,
                    "envio_s": sent,
                    "total_s": total,
                    "saida_mb": output.stat().st_size / 1024**2 if single else None,
                    "texto": _has_text(output) if single else "—",
                }
            )
    return results


def to_markdown(results: list[dict], printer: str) -> str:
    lines = [
        f"### Benchmark de impressão — {printer}",
        "",
        "| Cenário | Modo | Envio ao spooler | Até sair da fila | Arquivo gerado | Texto |",
        "|---|---|---:|---:|---:|:---:|",
    ]
    for r in results:
        out = f"{r['saida_mb']:.2f} MB" if r["saida_mb"] is not None else "—"
        lines.append(
            f"| {r['cenario']} | {r['modo']} | {r['envio_s']:.2f} s | {r['total_s']:.2f} s "
            f"| {out} | {r['texto']} |"
        )
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("impressora")
    parser.add_argument("--paginas", type=int, default=20)
    parser.add_argument("--saida", type=Path)
    args = parser.parse_args()
    if not gdi.vector_available():
        sys.exit("Impressão vetorial indisponível: rode no Windows.")
    text = to_markdown(run(args.impressora, args.paginas, args.saida), args.impressora)
    print(text)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as fh:
            fh.write(text + "\n")


if __name__ == "__main__":
    main()
