"""Janela principal (CustomTkinter)."""

from __future__ import annotations

import itertools
import logging
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from customtkinter.windows.widgets.appearance_mode import AppearanceModeTracker
from customtkinter.windows.widgets.scaling import ScalingTracker
from PIL import Image

from .. import APP_NAME, __version__
from ..backend import Backend
from ..files import Collector, Entry, FileKind
from ..jobs import JobEvent, PrintJob, Status
from ..paths import log_dir, resource_path
from ..printers import PrinterInfo
from ..settings import MAX_COPIES, Settings, load_settings, save_settings
from . import theme

log = logging.getLogger(__name__)

try:  # arrastar e soltar é opcional: se o tkdnd não carregar, o app funciona sem ele
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _Base = type("_Base", (ctk.CTk, TkinterDnD.DnDWrapper), {})
except Exception:  # noqa: BLE001
    TkinterDnD = None
    DND_FILES = None
    _Base = ctk.CTk

APPEARANCE_LABELS = {"system": "Sistema", "light": "Claro", "dark": "Escuro"}
KIND_ICONS = {
    FileKind.PDF: "PDF",
    FileKind.IMAGE: "Imagem",
    FileKind.WORD: "Word",
    FileKind.EXCEL: "Excel",
    FileKind.ARCHIVE: "Compactado",
    FileKind.UNSUPPORTED: "—",
}
STATUS_TEXT = {
    Status.PENDING: "●  Na fila",
    Status.PRINTING: "◌  Imprimindo…",
    Status.DONE: "✔  Impresso",
    Status.ERROR: "✖  Erro",
    Status.SKIPPED: "—  Ignorado",
    Status.CANCELLED: "■  Cancelado",
}
RETRYABLE = (Status.PENDING, Status.ERROR, Status.CANCELLED)
FILE_TYPES = [
    (
        "Arquivos suportados",
        "*.pdf *.png *.jpg *.jpeg *.jfif *.bmp *.gif *.tif *.tiff *.webp *.doc *.docx *.docm "
        "*.rtf *.odt *.txt *.xls *.xlsx *.xlsm *.xlsb *.csv *.ods *.zip *.rar *.7z *.tar *.tgz",
    ),
    ("Todos os arquivos", "*.*"),
]


@dataclass
class Item:
    entry: Entry
    status: Status
    detail: str = ""
    pages: int | None = None


class App(_Base):
    def __init__(
        self,
        backend: Backend,
        settings_file: Path | None = None,
        on_ready: Callable[[App], None] | None = None,
    ) -> None:
        self.backend = backend
        self.settings_file = settings_file
        self.settings: Settings = load_settings(settings_file)
        ctk.set_appearance_mode(self.settings.appearance)
        theme.apply()
        super().__init__()

        self.dnd_enabled = self._init_dnd()
        self.items: dict[str, Item] = {}
        self.printers: list[PrinterInfo] = []
        self.job: PrintJob | None = None
        self._job_iids: list[str] = []
        self._preparing = 0
        self._ids = itertools.count(1)
        self._ui_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._job_events: queue.Queue[JobEvent] = queue.Queue()
        self.workdir = Path(tempfile.mkdtemp(prefix="printer_cav_"))
        self.collector = Collector(self.workdir)

        self._setup_window()
        self._build_sidebar()
        self._build_main()
        self._apply_tree_style()
        AppearanceModeTracker.add(lambda _mode: self._apply_tree_style(), self)
        self._refresh_ui_state()
        self.load_printers()
        self.after(60, self._poll_queues)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        if on_ready is not None:
            self.after(300, lambda: on_ready(self))

    # ------------------------------------------------------------------ janela

    def _init_dnd(self) -> bool:
        if TkinterDnD is None:
            return False
        try:
            self.TkdndVersion = TkinterDnD._require(self)
            return True
        except Exception as exc:  # noqa: BLE001
            log.warning("Arrastar e soltar indisponível: %s", exc)
            return False

    def _setup_window(self) -> None:
        self.title(APP_NAME)
        self.minsize(860, 520)
        self._fit_to_screen(1040, 640)
        icon = resource_path("assets/favicon.ico")
        try:
            if os.name == "nt":
                import ctypes

                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CAVSUL.PrinterCAV")
                self.iconbitmap(str(icon))
            else:
                self._icon_photo = tk.PhotoImage(
                    master=self, file=str(resource_path("assets/logo.png"))
                )
                self.iconphoto(True, self._icon_photo)
        except Exception as exc:  # noqa: BLE001 - ícone nunca deve impedir o app de abrir
            log.warning("Ícone não carregado: %s", exc)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _fit_to_screen(self, width: int, height: int) -> None:
        """Tamanho inicial que cabe na tela (notebooks 1366x768 com escala de 125%) e centraliza."""
        try:
            scale = ScalingTracker.get_window_scaling(self)
        except Exception:  # noqa: BLE001
            scale = 1.0
        screen_w = self.winfo_screenwidth() / scale
        screen_h = self.winfo_screenheight() / scale
        w = int(min(width, max(860, screen_w * 0.92)))
        h = int(min(height, max(520, screen_h * 0.85)))
        x = int(max(0, (screen_w - w) / 2))
        y = int(max(0, (screen_h - h) / 3))
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------ barra lateral

    def _build_sidebar(self) -> None:
        bar = ctk.CTkFrame(self, width=290, corner_radius=0, fg_color=theme.SIDEBAR)
        bar.grid(row=0, column=0, sticky="nsw")
        bar.grid_propagate(False)
        bar.grid_columnconfigure(0, weight=1)
        self.sidebar = bar
        pad = {"padx": 22}

        logo = Image.open(resource_path("assets/logo.png"))
        self._logo = ctk.CTkImage(light_image=logo, dark_image=logo, size=(150, 84))
        ctk.CTkLabel(bar, image=self._logo, text="").grid(row=0, column=0, pady=(24, 6), **pad)
        ctk.CTkLabel(
            bar, text="Impressão em lote", font=theme.font(15, "bold"), text_color=theme.TEXT
        ).grid(row=1, column=0, **pad)

        theme.section(bar, "IMPRESSORA").grid(row=2, column=0, sticky="w", pady=(22, 6), **pad)
        row = ctk.CTkFrame(bar, fg_color="transparent")
        row.grid(row=3, column=0, sticky="ew", **pad)
        row.grid_columnconfigure(0, weight=1)
        self.printer_var = ctk.StringVar(value="Carregando…")
        self.printer_menu = ctk.CTkOptionMenu(
            row,
            variable=self.printer_var,
            values=["Carregando…"],
            command=self._on_printer_selected,
            dynamic_resizing=False,
            font=theme.font(13),
        )
        self.printer_menu.grid(row=0, column=0, sticky="ew")
        self.refresh_button = ctk.CTkButton(
            row, text="", image=theme.icon("refresh", 16), width=34, command=self.load_printers
        )
        self.refresh_button.grid(row=0, column=1, padx=(6, 0))
        theme.tooltip(self.refresh_button, "Atualizar lista de impressoras")
        self.printer_status = ctk.CTkLabel(bar, text="", font=theme.font(12), anchor="w")
        self.printer_status.grid(row=4, column=0, sticky="ew", pady=(4, 0), **pad)
        self.prefs_button = theme.secondary_button(
            bar, "Preferências da impressora", self.open_printer_preferences
        )
        self.prefs_button.grid(row=5, column=0, sticky="ew", pady=(8, 0), **pad)

        theme.section(bar, "CÓPIAS").grid(row=6, column=0, sticky="w", pady=(20, 6), **pad)
        stepper = ctk.CTkFrame(bar, fg_color="transparent")
        stepper.grid(row=7, column=0, sticky="w", **pad)
        self.copies_var = ctk.StringVar(value=str(self.settings.copies))
        self.copies_minus = ctk.CTkButton(
            stepper, text="−", width=34, command=lambda: self._step_copies(-1)
        )
        self.copies_entry = ctk.CTkEntry(
            stepper, textvariable=self.copies_var, width=56, justify="center", font=theme.font(13)
        )
        self.copies_plus = ctk.CTkButton(
            stepper, text="+", width=34, command=lambda: self._step_copies(1)
        )
        self.copies_minus.grid(row=0, column=0)
        self.copies_entry.grid(row=0, column=1, padx=6)
        self.copies_plus.grid(row=0, column=2)
        self.copies_entry.bind("<FocusOut>", lambda _e: self._normalize_copies())
        self.fast_pdf_switch = ctk.CTkSwitch(
            stepper, text="PDF rápido (vetorial)", font=theme.font(12),
            command=self._on_pdf_mode, progress_color=theme.ACCENT,
        )  # fmt: skip
        if self.settings.pdf_mode == "vector":
            self.fast_pdf_switch.select()
        self.fast_pdf_switch.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))
        theme.tooltip(
            self.fast_pdf_switch,
            "Envia o PDF como texto e vetores (como o Foxit): menos dados para a impressora.\n"
            "Desligue se alguma impressora imprimir PDFs com defeito.",
        )

        theme.section(bar, "ADICIONAR").grid(row=8, column=0, sticky="w", pady=(20, 6), **pad)
        self.add_files_button = ctk.CTkButton(
            bar, text="Arquivos…", command=self.choose_files, font=theme.font(13, "bold")
        )
        self.add_files_button.grid(row=9, column=0, sticky="ew", **pad)
        self.add_folder_button = ctk.CTkButton(
            bar, text="Pasta…", command=self.choose_folder, font=theme.font(13, "bold")
        )
        self.add_folder_button.grid(row=10, column=0, sticky="ew", pady=(8, 0), **pad)

        bar.grid_rowconfigure(11, weight=1)
        self.appearance = ctk.CTkSegmentedButton(
            bar,
            values=list(APPEARANCE_LABELS.values()),
            command=self._on_appearance,
            font=theme.font(12),
        )
        self.appearance.set(APPEARANCE_LABELS[self.settings.appearance])
        self.appearance.grid(row=12, column=0, sticky="ew", pady=(0, 10), **pad)
        footer = ctk.CTkFrame(bar, fg_color="transparent")
        footer.grid(row=13, column=0, sticky="ew", pady=(0, 16), **pad)
        ctk.CTkLabel(
            footer, text=f"v{__version__}", font=theme.font(11), text_color=theme.MUTED
        ).pack(side="left")
        logs = ctk.CTkLabel(
            footer,
            text="Abrir logs",
            font=theme.font(11, underline=True),
            text_color=theme.MUTED,
            cursor="hand2",
        )
        logs.pack(side="right")
        logs.bind("<Button-1>", lambda _e: open_path(log_dir()))

    # ------------------------------------------------------------------ área principal

    def _build_main(self) -> None:
        main = ctk.CTkFrame(self, corner_radius=0, fg_color=theme.BACKGROUND)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)
        self.main = main

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 10))
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header, text="Fila de impressão", font=theme.font(20, "bold"), text_color=theme.TEXT
        ).grid(row=0, column=0, sticky="w")
        self.count_label = ctk.CTkLabel(
            header, text="", font=theme.font(13), text_color=theme.MUTED
        )
        self.count_label.grid(row=0, column=1, sticky="w", padx=12)
        self.remove_button = theme.secondary_button(header, "Remover", self.remove_selected, 96)
        self.remove_button.grid(row=0, column=2, padx=(0, 8))
        self.clear_button = theme.secondary_button(header, "Limpar lista", self.clear_list, 110)
        self.clear_button.grid(row=0, column=3)

        card = ctk.CTkFrame(
            main, corner_radius=12, fg_color=theme.CARD, border_width=1, border_color=theme.BORDER
        )
        card.grid(row=1, column=0, sticky="nsew", padx=24)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)
        self.card = card

        columns = ("name", "origin", "kind", "pages", "status")
        self.tree = ttk.Treeview(
            card, columns=columns, show="headings", style="Printer.Treeview", selectmode="extended"
        )
        for col, title, width, stretch, anchor in (
            ("name", "Arquivo", 230, True, "w"),
            ("origin", "Origem", 130, True, "w"),
            ("kind", "Tipo", 80, False, "w"),
            ("pages", "Págs.", 54, False, "center"),
            ("status", "Status", 190, True, "w"),
        ):
            self.tree.heading(col, text=title, anchor=anchor)
            self.tree.column(
                col, width=self._px(width), minwidth=self._px(50), stretch=stretch, anchor=anchor
            )
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        scroll = ctk.CTkScrollbar(card, command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns", padx=(2, 6), pady=10)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", lambda _e: self.open_selected())
        self.tree.bind("<Delete>", lambda _e: self.remove_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_selected_detail())
        self.tree.bind("<Button-3>", self._context_menu)

        self.empty = ctk.CTkFrame(card, fg_color="transparent")
        ctk.CTkLabel(self.empty, text="", image=theme.icon("drop", 64)).pack(pady=(0, 6))
        drop_text = (
            "Arraste arquivos, pastas ou compactados para cá"
            if self.dnd_enabled
            else "Adicione arquivos, pastas ou compactados"
        )
        ctk.CTkLabel(
            self.empty, text=drop_text, font=theme.font(17, "bold"), text_color=theme.TEXT
        ).pack(pady=(4, 2))
        ctk.CTkLabel(
            self.empty,
            text="ou use os botões Arquivos… e Pasta… ao lado",
            font=theme.font(13),
            text_color=theme.MUTED,
        ).pack()
        ctk.CTkLabel(
            self.empty,
            text="PDF  •  Imagens  •  Word  •  Excel  •  CSV  •  ZIP / RAR / 7z",
            font=theme.font(12),
            text_color=theme.MUTED,
        ).pack(pady=(14, 0))

        footer = ctk.CTkFrame(main, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(12, 20))
        footer.grid_columnconfigure(0, weight=1)
        self.status_label = ctk.CTkLabel(
            footer,
            text="",
            font=theme.font(13),
            text_color=theme.MUTED,
            anchor="w",
            justify="left",
            wraplength=520,
        )
        self.status_label.grid(row=0, column=0, sticky="ew")
        self.progress = ctk.CTkProgressBar(footer, height=8)
        self.progress.set(0)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(6, 0), padx=(0, 16))
        self.print_button = ctk.CTkButton(
            footer,
            text="Imprimir",
            width=170,
            height=44,
            font=theme.font(15, "bold"),
            command=self.print_or_cancel,
        )
        self.print_button.grid(row=0, column=1, rowspan=2)

        if self.dnd_enabled:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Abrir arquivo", command=self.open_selected)
        self.menu.add_command(label="Remover da lista", command=self.remove_selected)

    def _apply_tree_style(self) -> None:
        dark = ctk.get_appearance_mode() == "Dark"
        colors = theme.tree_colors(dark)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Printer.Treeview",
            background=colors["bg"],
            fieldbackground=colors["bg"],
            foreground=colors["fg"],
            rowheight=self._px(32),
            borderwidth=0,
            font=theme.tk_font(10),
        )
        style.map(
            "Printer.Treeview",
            background=[("selected", colors["select_bg"])],
            foreground=[("selected", colors["select_fg"])],
        )
        style.configure(
            "Printer.Treeview.Heading",
            background=colors["bg"],
            foreground=colors["muted"],
            borderwidth=0,
            relief="flat",
            font=theme.tk_font(9, "bold"),
            padding=(6, 6),
        )
        style.map("Printer.Treeview.Heading", background=[("active", colors["bg"])])
        style.layout("Printer.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        for status, color in colors["status"].items():
            self.tree.tag_configure(status.name, foreground=color)
        self.menu.configure(
            bg=colors["bg"],
            fg=colors["fg"],
            activebackground=colors["select_bg"],
            activeforeground=colors["select_fg"],
            bd=0,
        )

    def _px(self, value: int) -> int:
        """Converte pixels lógicos para a escala do monitor (125%, 150%...); o ttk não escala."""
        try:
            scale = ScalingTracker.get_widget_scaling(self)
        except Exception:  # noqa: BLE001
            scale = 1.0
        return round(value * scale)

    # ------------------------------------------------------------------ impressoras

    def load_printers(self) -> None:
        self.refresh_button.configure(state="disabled")
        self.printer_status.configure(text="Procurando impressoras…", text_color=theme.MUTED)
        threading.Thread(target=self._load_printers_worker, name="printers", daemon=True).start()

    def _load_printers_worker(self) -> None:
        try:
            result: object = self.backend.list_printers()
        except Exception as exc:
            log.exception("Falha ao listar impressoras")
            result = exc
        self._ui_queue.put(("printers", result))

    def _set_printers(self, result: object) -> None:
        self.refresh_button.configure(state="normal")
        if isinstance(result, Exception):
            self.printers = []
            self.printer_menu.configure(values=["Nenhuma impressora"])
            self.printer_var.set("Nenhuma impressora")
            self.printer_status.configure(text=f"Erro ao listar: {result}", text_color=theme.ERROR)
            self._refresh_ui_state()
            return
        self.printers = list(result)  # type: ignore[arg-type]
        names = [p.name for p in self.printers]
        if not names:
            self.printer_menu.configure(values=["Nenhuma impressora"])
            self.printer_var.set("Nenhuma impressora")
            self.printer_status.configure(
                text="Nenhuma impressora instalada", text_color=theme.ERROR
            )
            self._refresh_ui_state()
            return
        self.printer_menu.configure(values=names)
        current = self.printer_var.get()
        default = next((p.name for p in self.printers if p.is_default), names[0])
        for candidate in (current, self.settings.printer, default):
            if candidate in names:
                self.printer_var.set(candidate)
                break
        self._on_printer_selected(self.printer_var.get())

    @property
    def selected_printer(self) -> PrinterInfo | None:
        name = self.printer_var.get()
        return next((p for p in self.printers if p.name == name), None)

    def _on_printer_selected(self, _name: str) -> None:
        printer = self.selected_printer
        if printer is None:
            return
        if printer.problem:
            text, color = f"●  {printer.problem.capitalize()}", theme.ERROR
        elif printer.is_default:
            text, color = "●  Disponível  •  padrão do Windows", theme.SUCCESS
        else:
            text, color = "●  Disponível", theme.SUCCESS
        self.printer_status.configure(text=text, text_color=color)
        self.settings.printer = printer.name
        self._refresh_ui_state()

    def open_printer_preferences(self) -> None:
        printer = self.selected_printer
        if printer is None:
            self._set_status("Selecione uma impressora primeiro.", theme.ERROR)
            return
        try:
            self.backend.open_preferences(printer.name)
        except Exception as exc:  # noqa: BLE001
            self._set_status(f"Não foi possível abrir as preferências: {exc}", theme.ERROR)

    # ------------------------------------------------------------------ cópias / tema

    @property
    def copies(self) -> int:
        try:
            value = int(self.copies_var.get())
        except ValueError:
            value = 1
        return min(max(value, 1), MAX_COPIES)

    def _normalize_copies(self) -> None:
        self.copies_var.set(str(self.copies))
        self.settings.copies = self.copies

    def _step_copies(self, delta: int) -> None:
        self.copies_var.set(str(min(max(self.copies + delta, 1), MAX_COPIES)))
        self.settings.copies = self.copies

    def _on_pdf_mode(self) -> None:
        self.settings.pdf_mode = "vector" if self.fast_pdf_switch.get() else "raster"

    def _on_appearance(self, label: str) -> None:
        mode = next(k for k, v in APPEARANCE_LABELS.items() if v == label)
        self.settings.appearance = mode
        ctk.set_appearance_mode(mode)
        self._apply_tree_style()

    # ------------------------------------------------------------------ adicionar arquivos

    def choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self, title="Selecione os arquivos", filetypes=FILE_TYPES
        )
        if paths:
            self.add_paths(paths)

    def choose_folder(self) -> None:
        path = filedialog.askdirectory(parent=self, title="Selecione a pasta", mustexist=True)
        if path:
            self.add_paths([path])

    def _on_drop(self, event) -> None:
        self.add_paths(self.tk.splitlist(event.data))

    def add_paths(self, paths: Iterable[str]) -> None:
        paths = [p for p in paths if p]
        if not paths or self._is_printing:
            return
        self._preparing += 1
        self._refresh_ui_state()
        threading.Thread(
            target=self._collect_worker, args=(paths,), name="collect", daemon=True
        ).start()

    def _collect_worker(self, paths: list[str]) -> None:
        try:
            result: object = self.collector.collect(paths)
        except Exception as exc:
            log.exception("Falha ao preparar arquivos")
            result = exc
        self._ui_queue.put(("collected", result))

    def _add_entries(self, result: object) -> None:
        self._preparing -= 1
        if isinstance(result, Exception):
            self._set_status(f"Erro ao preparar arquivos: {result}", theme.ERROR)
        else:
            known = {item.entry.path for item in self.items.values()}
            duplicates = 0
            for entry in result:  # type: ignore[union-attr]
                if entry.kind is FileKind.ARCHIVE and entry.error is None:
                    continue
                if entry.path in known:
                    duplicates += 1
                    continue
                known.add(entry.path)
                status = Status.PENDING if entry.printable else Status.SKIPPED
                detail = entry.error or ("" if entry.printable else "não suportado")
                iid = str(next(self._ids))
                self.items[iid] = Item(entry, status, detail)
                self.tree.insert(
                    "", "end", iid=iid, values=self._row(self.items[iid]), tags=(status.name,)
                )
            self._set_status(
                f"{duplicates} arquivo(s) já estava(m) na lista e não foram repetidos."
                if duplicates
                else ""
            )
        self._refresh_ui_state()

    @staticmethod
    def _row(item: Item) -> tuple[str, ...]:
        status = STATUS_TEXT[item.status]
        if item.detail and item.status in (Status.ERROR, Status.SKIPPED):
            status = f"{status}: {item.detail}"
        pages = "" if item.pages is None else str(item.pages)
        return (item.entry.display, item.entry.origin, KIND_ICONS[item.entry.kind], pages, status)

    def _update_item(self, iid: str) -> None:
        if self.tree.exists(iid):
            item = self.items[iid]
            self.tree.item(iid, values=self._row(item), tags=(item.status.name,))

    # ------------------------------------------------------------------ lista

    def remove_selected(self) -> None:
        if self._is_printing:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.items.pop(iid, None)
        self._refresh_ui_state()

    def clear_list(self) -> None:
        if self._is_printing or self._preparing:
            return
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        shutil.rmtree(self.workdir, ignore_errors=True)
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.progress.set(0)
        self._set_status("")
        self._refresh_ui_state()

    def open_selected(self) -> None:
        for iid in self.tree.selection()[:5]:
            path = self.items[iid].entry.path
            if path.exists():
                open_path(path)

    def _show_selected_detail(self) -> None:
        selection = self.tree.selection()
        if len(selection) == 1 and self.items[selection[0]].detail and not self._is_printing:
            item = self.items[selection[0]]
            self._set_status(
                f"{item.entry.display}: {item.detail}",
                theme.ERROR if item.status is Status.ERROR else theme.MUTED,
            )

    def _context_menu(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.menu.tk_popup(event.x_root, event.y_root)

    # ------------------------------------------------------------------ impressão

    @property
    def _is_printing(self) -> bool:
        return self.job is not None

    def print_or_cancel(self) -> None:
        if self._is_printing:
            self.cancel_print()
        else:
            self.start_print()

    def start_print(self, confirm: bool = True) -> bool:
        printer = self.selected_printer
        if printer is None:
            self._set_status("Selecione uma impressora.", theme.ERROR)
            return False
        order = list(self.tree.get_children())
        targets = [i for i in order if self.items[i].status in RETRYABLE]
        if not targets:
            printed = [i for i in order if self.items[i].status is Status.DONE]
            if not printed:
                self._set_status("Nada para imprimir: adicione arquivos suportados.", theme.ERROR)
                return False
            if confirm and not messagebox.askyesno(
                APP_NAME,
                "Todos os arquivos já foram impressos.\n\nImprimir tudo novamente?",
                parent=self,
            ):
                return False
            targets = printed
        if confirm:
            problem = self._safe_problem(printer.name)
            if problem and not messagebox.askyesno(
                APP_NAME,
                f"A impressora “{printer.name}” está {problem}.\n\nEnviar mesmo assim? "
                "Os documentos ficarão na fila do Windows até ela voltar.",
                icon="warning",
                parent=self,
            ):
                return False
        self._normalize_copies()
        save_settings(self.settings, self.settings_file)

        for iid in targets:
            item = self.items[iid]
            item.status, item.detail, item.pages = Status.PENDING, "", None
            self._update_item(iid)
        self._job_iids = targets
        self.job = PrintJob(
            [(n, self.items[iid].entry) for n, iid in enumerate(targets)],
            printer.name,
            self.copies,
            self._job_events,
            print_fn=self.backend.print_fn,
            session_factory=self.backend.session_factory,
            pdf_mode=self.settings.pdf_mode,
        )
        self.progress.set(0)
        self._set_status(f"Enviando para {printer.name}…")
        self._refresh_ui_state()
        self.job.start()
        return True

    def _safe_problem(self, name: str) -> str | None:
        try:
            return self.backend.printer_problem(name)
        except Exception as exc:  # noqa: BLE001
            log.warning("Status da impressora indisponível: %s", exc)
            return None

    def cancel_print(self) -> None:
        if self.job is not None:
            self.job.cancel()
            self.print_button.configure(state="disabled", text="Cancelando…")

    def _on_job_event(self, event: JobEvent) -> None:
        if event.index is None:
            self._finish_job()
            return
        iid = self._job_iids[event.index]
        item = self.items.get(iid)
        if item is None:
            return
        item.status, item.detail, item.pages = event.status, event.detail, event.pages
        self._update_item(iid)
        if event.status is Status.PRINTING:
            self.tree.see(iid)
            self._set_status(
                f"Imprimindo {event.index + 1} de {len(self._job_iids)}: {item.entry.display}"
            )
        else:
            self.progress.set((event.index + 1) / len(self._job_iids))

    def _finish_job(self) -> None:
        summary = self.job.summary if self.job else None
        self.job = None
        self._job_iids = []
        self.progress.set(1 if summary and not summary.cancelled else self.progress.get())
        if summary is not None:
            parts = [f"{summary.done} impresso(s)"]
            elapsed = f" em {_duration(summary.seconds)}" if summary.seconds >= 1 else ""
            if summary.pages:
                parts[0] += f" • {summary.pages} página(s)"
            if summary.errors:
                parts.append(f"{summary.errors} com erro")
            if summary.skipped:
                parts.append(f"{summary.skipped} ignorado(s)")
            if summary.cancelled:
                parts.append(f"{summary.cancelled} cancelado(s)")
            color = theme.ERROR if summary.errors else theme.SUCCESS
            self._set_status(f"Concluído{elapsed}: " + ", ".join(parts), color)
            if summary.errors or summary.cancelled:
                self.bell()
        self._refresh_ui_state()

    # ------------------------------------------------------------------ estado e eventos

    def _poll_queues(self) -> None:
        try:
            while True:
                kind, payload = self._ui_queue.get_nowait()
                if kind == "printers":
                    self._set_printers(payload)
                elif kind == "collected":
                    self._add_entries(payload)
        except queue.Empty:
            pass
        try:
            for _ in range(200):
                self._on_job_event(self._job_events.get_nowait())
        except queue.Empty:
            pass
        self.after(60, self._poll_queues)

    def _set_status(self, text: str, color: str | tuple[str, str] = theme.MUTED) -> None:
        self.status_label.configure(text=text, text_color=color)

    def _refresh_ui_state(self) -> None:
        printing = self._is_printing
        has_items = bool(self.items)
        printable = sum(1 for i in self.items.values() if i.status is not Status.SKIPPED)
        skipped = len(self.items) - printable
        idle = "disabled" if printing else "normal"

        for widget in (
            self.add_files_button,
            self.add_folder_button,
            self.copies_minus,
            self.copies_plus,
            self.copies_entry,
            self.printer_menu,
            self.prefs_button,
            self.appearance,
            self.fast_pdf_switch,
        ):
            widget.configure(state=idle)
        busy = printing or self._preparing > 0
        self.clear_button.configure(state="disabled" if busy or not has_items else "normal")
        self.remove_button.configure(state="disabled" if printing or not has_items else "normal")

        if printing:
            self.print_button.configure(
                text="Cancelar",
                state="normal",
                fg_color=theme.DANGER,
                hover_color=theme.DANGER_HOVER,
            )
        else:
            can_print = printable > 0 and self.selected_printer is not None and not self._preparing
            self.print_button.configure(
                text="Imprimir",
                state="normal" if can_print else "disabled",
                fg_color=theme.ACCENT,
                hover_color=theme.ACCENT_HOVER,
            )

        if self._preparing:
            self.progress.configure(mode="indeterminate")
            self.progress.start()
            self._set_status("Preparando arquivos…")
        elif str(self.progress.cget("mode")) == "indeterminate":
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(0)

        if has_items:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.5, anchor="center")
        text = f"{printable} arquivo(s)" if has_items else ""
        if skipped:
            text += f"  •  {skipped} ignorado(s)"
        self.count_label.configure(text=text)

    def on_close(self) -> None:
        if self._is_printing:
            if not messagebox.askyesno(
                APP_NAME, "Há uma impressão em andamento.\n\nCancelar e sair?", parent=self
            ):
                return
            self.job.cancel()
            self.job.join(timeout=10)
        self._normalize_copies()
        save_settings(self.settings, self.settings_file)
        shutil.rmtree(self.workdir, ignore_errors=True)
        self.destroy()


def _duration(seconds: float) -> str:
    minutes, secs = divmod(round(seconds), 60)
    return f"{minutes} min {secs:02d} s" if minutes else f"{secs} s"


def open_path(path: Path) -> None:
    """Abre um arquivo/pasta com o programa padrão do sistema."""
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, str(path)])
    except OSError as exc:
        log.warning("Não foi possível abrir %s: %s", path, exc)
