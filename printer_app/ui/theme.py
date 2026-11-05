"""Cores, fontes e componentes visuais (identidade azul da CAV SUL, modo claro/escuro)."""

from __future__ import annotations

import os
import tkinter as tk

import customtkinter as ctk
from PIL import Image, ImageDraw

from ..jobs import Status

# Pares (modo claro, modo escuro)
ACCENT = ("#2F5DA8", "#3D6FC4")
ACCENT_HOVER = ("#244A87", "#335EA8")
DANGER = ("#C62828", "#D84343")
DANGER_HOVER = ("#A61F1F", "#B83636")
SUCCESS = ("#2E7D32", "#5CC269")
ERROR = ("#C62828", "#FF7B7B")
TEXT = ("#1C2433", "#E8ECF3")
MUTED = ("#5D6778", "#9AA4B5")
BACKGROUND = ("#F3F5F9", "#161A22")
SIDEBAR = ("#FFFFFF", "#1E2430")
CARD = ("#FFFFFF", "#1E2430")
BORDER = ("#E1E6EF", "#2C3442")
SECONDARY_HOVER = ("#E8EEF8", "#2A3343")
ICON_LIGHT, ICON_DARK = "#FFFFFF", "#FFFFFF"

_FAMILY = "Segoe UI" if os.name == "nt" else "DejaVu Sans"


def apply() -> None:
    """Aplica a paleta azul da marca sobre o tema padrão do CustomTkinter."""
    ctk.set_default_color_theme("blue")
    theme = ctk.ThemeManager.theme
    theme["CTkButton"].update(
        fg_color=list(ACCENT), hover_color=list(ACCENT_HOVER), corner_radius=8
    )
    theme["CTkOptionMenu"].update(
        fg_color=list(ACCENT),
        button_color=list(ACCENT_HOVER),
        button_hover_color=list(ACCENT_HOVER),
        corner_radius=8,
    )
    theme["CTkProgressBar"]["progress_color"] = list(ACCENT)
    theme["CTkSegmentedButton"].update(
        selected_color=list(ACCENT), selected_hover_color=list(ACCENT_HOVER)
    )


def font(size: int, weight: str = "normal", underline: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(family=_FAMILY, size=size, weight=weight, underline=underline)


def tk_font(size: int, weight: str = "normal") -> tuple:
    return (_FAMILY, size, weight) if weight != "normal" else (_FAMILY, size)


def section(master, text: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(master, text=text, font=font(11, "bold"), text_color=MUTED)


def secondary_button(master, text: str, command, width: int | None = None) -> ctk.CTkButton:
    kwargs = {"width": width} if width else {}
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        fg_color="transparent",
        border_width=1,
        border_color=BORDER,
        text_color=TEXT,
        hover_color=SECONDARY_HOVER,
        font=font(13),
        **kwargs,
    )


def tree_colors(dark: bool) -> dict:
    i = 1 if dark else 0
    return {
        "bg": CARD[i],
        "fg": TEXT[i],
        "muted": MUTED[i],
        "select_bg": ("#DCE6F7", "#2D4A7A")[i],
        "select_fg": TEXT[i],
        "status": {
            Status.PENDING: TEXT[i],
            Status.PRINTING: ACCENT[i],
            Status.DONE: SUCCESS[i],
            Status.ERROR: ERROR[i],
            Status.SKIPPED: MUTED[i],
            Status.CANCELLED: MUTED[i],
        },
    }


class tooltip:  # noqa: N801 - usado como função
    """Dica simples ao passar o mouse."""

    def __init__(self, widget, text: str) -> None:
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _event=None) -> None:
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip, text=self.text, bg="#1C2433", fg="#FFFFFF", padx=8, pady=4, font=tk_font(9)
        ).pack()

    def _hide(self, _event=None) -> None:
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


# --- ícones desenhados em tempo de execução (nítidos em qualquer escala/tema) ---------


def _canvas(size: int = 256) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)


def _refresh_image(color: str) -> Image.Image:
    img, draw = _canvas()
    draw.arc((44, 44, 212, 212), start=-20, end=272, fill=color, width=26)
    # ponta da seta no fim do arco (-80°), apontando no sentido horário
    draw.polygon([(132, 6), (200, 52), (128, 92)], fill=color)
    return img


def _drop_image(color: str) -> Image.Image:
    img, draw = _canvas()
    draw.rounded_rectangle((28, 150, 228, 236), radius=26, outline=color, width=18)
    draw.rectangle((70, 140, 186, 168), fill=(0, 0, 0, 0))
    draw.rectangle((116, 20, 140, 140), fill=color)
    draw.polygon([(78, 108), (178, 108), (128, 170)], fill=color)
    return img


def icon(name: str, size: int) -> ctk.CTkImage:
    draw = {"refresh": _refresh_image, "drop": _drop_image}[name]
    light, dark = (ICON_LIGHT, ICON_DARK) if name == "refresh" else ACCENT
    return ctk.CTkImage(light_image=draw(light), dark_image=draw(dark), size=(size, size))
