"""Preparação de imagens para impressão."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from PIL import Image, ImageOps, ImageSequence


def to_printable(img: Image.Image) -> Image.Image:
    """Aplica a rotação EXIF (fotos de celular) e troca transparência por fundo branco."""
    img = ImageOps.exif_transpose(img)
    has_alpha = img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and "transparency" in img.info)
    if has_alpha:
        rgba = img.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        return Image.alpha_composite(background, rgba).convert("RGB")
    return img.convert("RGB")


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
