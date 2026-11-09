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


def _job_sizes(printer: str) -> int:
    """Maior tamanho de spool visto entre os trabalhos na fila (bytes)."""
    import win32print

    handle = win32print.OpenPrinter(printer, {"DesiredAccess": 0x8})
    try:
        jobs = win32print.EnumJobs(handle, 0, 999, 2)
    finally:
        win32print.ClosePrinter(handle)
    return max((job.get("Size", 0) for job in jobs), default=0)


def _wait_queue(printer: str, output: Path | None, timeout: float = 300) -> int:
    from printer_app.printers import pending_jobs

    deadline = time.monotonic() + timeout
    spool = 0
    last_size = -1
    while time.monotonic() < deadline:
        spool = max(spool, _job_sizes(printer))
        if pending_jobs(printer) == 0:
            if output is None:
                return spool
            size = output.stat().st_size if output.exists() else -1
            if size > 0 and size == last_size:
                return spool
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


def run(printer: str, pages: int, output: Path | None) -> list[dict]:
    with tempfile.TemporaryDirectory() as tmp:
        doc = make_text_pdf(Path(tmp) / "benchmark.pdf", pages=pages)
        modes = {
            "raster (v2.0.0)": lambda: gdi.print_pages(
                printer, "benchmark-raster", lambda dpi: pdf.pdf_pages(doc, dpi)
            ),
            "vetorial (novo)": lambda: gdi.print_pdf_vector(printer, doc, "benchmark-vetor"),
        }
        results = []
        for name, send in modes.items():
            if output is not None and output.exists():
                output.unlink()
            start = time.perf_counter()
            printed = send()
            sent = time.perf_counter() - start
            spool = _wait_queue(printer, output)
            total = time.perf_counter() - start
            results.append(
                {
                    "modo": name,
                    "paginas": printed,
                    "envio_s": sent,
                    "total_s": total,
                    "spool_mb": spool / 1024**2,
                    "saida_mb": output.stat().st_size / 1024**2 if output else None,
                    "texto": _has_text(output),
                }
            )
    return results


def to_markdown(results: list[dict], pages: int) -> str:
    lines = [
        f"### Benchmark de impressão — PDF de texto com {pages} páginas",
        "",
        "| Modo | Envio ao spooler | Até sair da fila | Spool | Arquivo gerado | Texto na saída |",
        "|---|---:|---:|---:|---:|:---:|",
    ]
    for r in results:
        out = f"{r['saida_mb']:.2f} MB" if r["saida_mb"] is not None else "—"
        lines.append(
            f"| {r['modo']} | {r['envio_s']:.2f} s | {r['total_s']:.2f} s "
            f"| {r['spool_mb']:.2f} MB | {out} | {r['texto']} |"
        )
    if len(results) == 2 and results[1]["total_s"] > 0:
        lines.append("")
        lines.append(
            f"**Vetorial {results[0]['total_s'] / results[1]['total_s']:.1f}x mais rápido "
            f"no tempo total.**"
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
    text = to_markdown(run(args.impressora, args.paginas, args.saida), args.paginas)
    print(text)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as fh:
            fh.write(text + "\n")


if __name__ == "__main__":
    main()
