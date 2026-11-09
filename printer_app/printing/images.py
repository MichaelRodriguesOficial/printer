"""Preparação de imagens para impressão."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from PIL import Image, ImageOps, ImageSequence


def to_printable(img: Image.Image) -> Image.Image:
    """Aplica a rotação EXIF (fotos de celular) e troca transparência por fundo branco.

    Imagens em tons de cinza (scanner P&B, documentos) continuam em 8 bits: um terço dos
    dados de uma imagem colorida no spool.
    """
    img = ImageOps.exif_transpose(img)
    grayscale = img.mode in ("1", "L", "LA", "I", "I;16", "F")
    has_alpha = img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info)
    if has_alpha:
        rgba = img.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        flat = Image.alpha_composite(background, rgba)
        return flat.convert("L" if grayscale else "RGB")
    if grayscale:
        return img.convert("L") if img.mode in ("1", "L") else _to_8bit_gray(img)
    return img.convert("RGB")


def _to_8bit_gray(img: Image.Image) -> Image.Image:
    """Modos de 16/32 bits (TIFF de scanner) normalizados para 0-255."""
    lo, hi = img.getextrema()
    if hi == lo:
        return Image.new("L", img.size, 255 if hi else 0)
    return img.point(lambda v: (v - lo) * 255 / (hi - lo)).convert("L")


def image_pages(path: Path) -> Iterator[Image.Image]:
    """Páginas de uma imagem: todas as páginas de um TIFF; só o 1º quadro de GIF/WebP animado."""
    with Image.open(path) as img:
        multipage_tiff = img.format == "TIFF" and getattr(img, "n_frames", 1) > 1
        frames = ImageSequence.Iterator(img) if multipage_tiff else [img]
        for frame in frames:
            yield to_printable(frame.copy())


def count_pages(path: Path) -> int:
    with Image.open(path) as img:
        return getattr(img, "n_frames", 1) if img.format == "TIFF" else 1
