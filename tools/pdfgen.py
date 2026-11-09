"""Gera PDFs de texto simples (sem dependências) para testes e benchmark de impressão."""

from __future__ import annotations

from pathlib import Path


def _escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def make_text_pdf(path: Path, pages: int = 1, lines: int = 45, title: str = "Pedido") -> Path:
    """PDF A4 com ``lines`` linhas de texto real (Helvetica) por página."""
    objects: list[str] = []
    page_ids = [4 + 2 * i for i in range(pages)]
    objects.append("<< /Type /Catalog /Pages 2 0 R >>")
    kids = " ".join(f"{pid} 0 R" for pid in page_ids)
    objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {pages} >>")
    objects.append(
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    )
    for n in range(pages):
        body = ["BT", "/F1 11 Tf", "14 TL", "56 790 Td"]
        body.append(f"({_escape(f'{title} {n + 1:03d} - Centro de Apoio a Vendas')}) Tj")
        for i in range(lines - 1):
            text = (
                f"Item {i + 1:02d}  Consultora {n * lines + i:05d}  Valor R$ {(i + 1) * 37.5:,.2f}"
            )
            body.append(f"T* ({_escape(text)}) Tj")
        body.append("ET")
        # Um retângulo vetorial simples por página (cabeçalho).
        body.append("0.18 0.36 0.66 rg 40 805 515 22 re f")
        stream = "\n".join(body)
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 3 0 R >> >> /Contents {page_ids[n] + 1} 0 R >>"
        )
        objects.append(
            f"<< /Length {len(stream.encode('latin-1'))} >>\nstream\n{stream}\nendstream"
        )

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for number, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{number} 0 obj\n{obj}\nendobj\n".encode("latin-1")
    xref = len(out)
    out += f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode()
    for offset in offsets:
        out += f"{offset:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    )
    path.write_bytes(bytes(out))
    return path
