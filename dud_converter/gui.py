"""Tkinter desktop window for the converter (Albanian / English).

Laid out as numbered steps for people who are not Excel experts: pick the
file, check the columns (each one shows whether it was found), answer the
questions, create the XML. The result opens in its own window.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import core
from .i18n import LANGUAGES, t

# Palette
PRIMARY = "#1F4E79"
PRIMARY_DARK = "#163A5A"
PRIMARY_LIGHT = "#DCE8F4"
BG = "#EEF2F6"
CARD = "#FFFFFF"
BORDER = "#D5DCE4"
TEXT = "#1E2A36"
MUTED = "#5F6B78"
OK = "#2E7D32"
BAD = "#C62828"
WARN_BG = "#FFF4E0"
WARN_TEXT = "#8A5300"
GO = "#2E7D32"
GO_DARK = "#24622A"

FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"


def _settings_path() -> Path:
    base = os.environ.get("APPDATA") if sys.platform == "win32" else None
    base = Path(base) if base else Path.home() / ".config"
    return base / "DUDConverter" / "settings.json"


def load_settings() -> dict:
    try:
        return json.loads(_settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(data: dict) -> None:
    try:
        path = _settings_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass  # settings are a convenience; never block the conversion


def open_folder(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa: S606 - opens Explorer on the user's own folder
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError:
        pass


class ScrollArea(tk.Frame):
    """A vertically scrolling container; put widgets in `.inner`."""

    def __init__(self, master, **kw):
        super().__init__(master, bg=BG, **kw)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=BG)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=bar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._wheel, add="+")

    def _wheel(self, event):
        if isinstance(event.widget, str) or event.widget.winfo_class() == "Listbox":
            return  # an open combobox list scrolls itself, not the page
        if event.num == 4:
            steps = -3
        elif event.num == 5:
            steps = 3
        else:
            steps = int(-event.delta / 120) * 3
        self.canvas.yview_scroll(steps, "units")

    def scroll_only(self, widget):
        """Make the wheel over `widget` scroll the page instead of changing its value.

        A ttk.Combobox steps through its list on every wheel or two-finger
        swipe, which silently changed the chosen columns while scrolling.
        """
        def wheel(event):
            self._wheel(event)
            return "break"
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(seq, wheel)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        self.lang = self.settings.get("language", "sq")
        if self.lang not in LANGUAGES:
            self.lang = "sq"

        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar(value=self.settings.get("output_folder", ""))
        # No default: the user must answer. "other" means the typed code is used.
        self.precision_var = tk.StringVar(value="")
        self.extended_var = tk.StringVar(value="")
        self.precision_other = tk.StringVar(value="")
        self.precision_auto = tk.BooleanVar(value=True)  # Precision_1 from TARIK per item
        self.extended_other = tk.StringVar(value="")
        saved_cols = self.settings.get("column_names", {})
        self.column_vars = {
            key: tk.StringVar(value=saved_cols.get(key, core.DEFAULT_COLUMN_NAMES[key]))
            for key, _ in core.FIELDS
        }
        self.indicators_var = tk.StringVar(
            value=self.settings.get("positive_indicators", ", ".join(core.DEFAULT_POSITIVE_INDICATORS))
        )
        self.country_vars = {}  # EU origin code -> BooleanVar, per invoice
        self.constants = dict(core.DEFAULT_CONSTANTS)
        self.constants.update(self.settings.get("constants", {}))
        self.sheet_info = None
        self._inspect_job = None

        self.configure(bg=BG)
        self.minsize(720, 560)
        self.geometry("900x780")
        self._setup_styles()
        for var in self.column_vars.values():
            var.trace_add("write", lambda *_: self._schedule_inspect())
        self._build()

    # ------------------------------------------------------------------ styles

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=(FONT, 10), background=CARD, foreground=TEXT)
        style.configure("TCombobox", padding=4)
        style.configure("TEntry", padding=4)
        style.configure("Choice.TRadiobutton", font=(FONT, 12, "bold"), padding=(18, 8),
                        background=PRIMARY_LIGHT, foreground=PRIMARY, indicatorthickness=0)
        style.map("Choice.TRadiobutton",
                  background=[("selected", PRIMARY), ("active", "#C7DAEE")],
                  foreground=[("selected", "#FFFFFF")])
        style.configure("Other.TRadiobutton", font=(FONT, 11), padding=(10, 8), background=CARD)
        style.configure("Code.TEntry", padding=6)
        style.configure("TCheckbutton", background=CARD, font=(FONT, 11))
        style.configure("Big.TButton", font=(FONT, 11, "bold"), padding=(16, 8),
                        background=PRIMARY, foreground="#FFFFFF", bordercolor=PRIMARY)
        style.map("Big.TButton", background=[("active", PRIMARY_DARK)])
        style.configure("Go.TButton", font=(FONT, 13, "bold"), padding=(28, 10),
                        background=GO, foreground="#FFFFFF", bordercolor=GO)
        style.map("Go.TButton", background=[("active", GO_DARK)])
        style.configure("Link.TButton", font=(FONT, 10), padding=(10, 4),
                        background=CARD, foreground=PRIMARY, bordercolor=BORDER)
        style.configure("Lang.TButton", font=(FONT, 10, "bold"), padding=(10, 3),
                        background=PRIMARY_DARK, foreground="#C9D7E6", bordercolor=PRIMARY_DARK)
        style.configure("LangOn.TButton", font=(FONT, 10, "bold"), padding=(10, 3),
                        background="#FFFFFF", foreground=PRIMARY, bordercolor="#FFFFFF")

    # ------------------------------------------------------------------ helpers

    def tr(self, key, **params):
        return t(self.lang, key, **params)

    def _label(self, parent, text, size=10, bold=False, color=TEXT, bg=CARD, **kw):
        return tk.Label(parent, text=text, font=(FONT, size, "bold" if bold else "normal"),
                        fg=color, bg=bg, anchor="w", justify="left", **kw)

    def _card(self, number, title, subtitle=""):
        outer = tk.Frame(self.body.inner, bg=BORDER)
        outer.pack(fill="x", padx=18, pady=(14, 0))
        card = tk.Frame(outer, bg=CARD, padx=18, pady=14)
        card.pack(fill="both", padx=1, pady=1)
        head = tk.Frame(card, bg=CARD)
        head.pack(fill="x")
        tk.Label(head, text=str(number), font=(FONT, 13, "bold"), fg="#FFFFFF", bg=PRIMARY,
                 width=2, pady=2).pack(side="left", padx=(0, 10))
        self._label(head, title, size=13, bold=True).pack(side="left")
        if subtitle:
            self._label(card, subtitle, color=MUTED, wraplength=780).pack(fill="x", pady=(6, 0))
        content = tk.Frame(card, bg=CARD)
        content.pack(fill="x", pady=(10, 0))
        return content

    # ------------------------------------------------------------------ layout

    def _build(self):
        """(Re)build the whole window; called again when the language changes."""
        for child in self.winfo_children():
            if not isinstance(child, tk.Toplevel):
                child.destroy()
        self.title(self.tr("title"))

        # Header bar
        header = tk.Frame(self, bg=PRIMARY, padx=20, pady=14)
        header.pack(fill="x")
        titles = tk.Frame(header, bg=PRIMARY)
        titles.pack(side="left", fill="x", expand=True)
        self._label(titles, self.tr("title"), size=16, bold=True, color="#FFFFFF", bg=PRIMARY).pack(anchor="w")
        self._label(titles, self.tr("subtitle"), color="#C9D7E6", bg=PRIMARY).pack(anchor="w")
        langs = tk.Frame(header, bg=PRIMARY)
        langs.pack(side="right")
        for code, name in LANGUAGES.items():
            ttk.Button(langs, text=name, style="LangOn.TButton" if code == self.lang else "Lang.TButton",
                       command=lambda c=code: self._set_language(c)).pack(side="left", padx=2)

        # Footer with the main action, always visible
        footer = tk.Frame(self, bg=CARD, padx=20, pady=12, highlightthickness=1, highlightbackground=BORDER)
        footer.pack(side="bottom", fill="x")
        left = tk.Frame(footer, bg=CARD)
        left.pack(side="left", fill="x", expand=True)
        self._label(left, self.tr("save_hint"), color=MUTED, wraplength=520).pack(anchor="w")
        links = tk.Frame(left, bg=CARD)
        links.pack(anchor="w", pady=(4, 0))
        ttk.Button(links, text="⚙ " + self.tr("constants"), style="Link.TButton",
                   command=self._edit_constants).pack(side="left")
        ttk.Button(footer, text=self.tr("create"), style="Go.TButton", command=self._create).pack(side="right")

        self.body = ScrollArea(self)
        self.body.pack(fill="both", expand=True)
        self._build_step_file()
        self._build_step_columns()
        self._build_step_questions()
        self._build_step_create()
        tk.Frame(self.body.inner, bg=BG, height=16).pack()
        self._refresh()

    def _build_step_file(self):
        c = self._card(1, self.tr("step1_title"), self.tr("step1_hint"))
        row = tk.Frame(c, bg=CARD)
        row.pack(fill="x")
        ttk.Button(row, text="📂  " + self.tr("choose_file"), style="Big.TButton",
                   command=self._browse_excel).pack(side="left")
        info = tk.Frame(row, bg=CARD)
        info.pack(side="left", fill="x", expand=True, padx=14)
        self.file_name_label = self._label(info, "", size=11, bold=True)
        self.file_name_label.pack(anchor="w")
        self.file_status_label = self._label(info, "", color=MUTED, wraplength=560)
        self.file_status_label.pack(anchor="w")

    def _build_step_columns(self):
        c = self._card(2, self.tr("step2_title"), self.tr("step2_hint"))
        c.columnconfigure(1, weight=1)
        self.column_boxes, self.column_status = {}, {}
        for i, (key, _) in enumerate(core.FIELDS):
            self._label(c, self.tr("field_" + key)).grid(row=i, column=0, sticky="w", pady=3, padx=(0, 12))
            box = ttk.Combobox(c, textvariable=self.column_vars[key], width=34)
            box.grid(row=i, column=1, sticky="ew", pady=3)
            self.body.scroll_only(box)
            box.bind("<<ComboboxSelected>>", self._column_picked)
            status = self._label(c, "", size=10, bold=True)
            status.grid(row=i, column=2, sticky="w", padx=(10, 0))
            self.column_boxes[key], self.column_status[key] = box, status

    def _column_picked(self, event):
        # Leave the box once a column is picked, so nothing can change it by accident.
        event.widget.selection_clear()
        self.focus_set()

    def _build_step_questions(self):
        c = self._card(3, self.tr("step3_title"), self.tr("step3_hint"))
        for label_key, hint_key, var, other, choices, digits in (
            ("q_precision", "q_code_hint", self.precision_var, self.precision_other,
             core.PRECISION_1_CHOICES, core.PRECISION_1_DIGITS),
            ("q_extended", "q_code_hint", self.extended_var, self.extended_other,
             core.EXTENDED_PROCEDURE_CHOICES, core.EXTENDED_PROCEDURE_DIGITS),
        ):
            row = tk.Frame(c, bg=CARD)
            row.pack(fill="x", pady=(0, 12))
            self._label(row, self.tr(label_key), size=11, bold=True).pack(anchor="w")
            if var is self.precision_var:
                ttk.Checkbutton(row, text=self.tr("q_precision_auto"), variable=self.precision_auto).pack(
                    anchor="w", pady=(4, 0))
                self._label(row, self.tr("q_precision_auto_hint"), color=MUTED, wraplength=780).pack(
                    anchor="w", pady=(2, 0))
            self._label(row, self.tr(hint_key, digits=digits), color=MUTED, wraplength=780).pack(
                anchor="w", pady=(2, 6))
            buttons = tk.Frame(row, bg=CARD)
            buttons.pack(anchor="w")
            for v in choices:
                ttk.Radiobutton(buttons, text=v, value=v, variable=var,
                                style="Choice.TRadiobutton").pack(side="left", padx=(0, 6))
            ttk.Radiobutton(buttons, text=self.tr("other_code"), value="other", variable=var,
                            style="Other.TRadiobutton").pack(side="left", padx=(10, 2))
            check = (self.register(lambda text, n=digits: len(text) <= n and (text == "" or text.isdigit())), "%P")
            entry = ttk.Entry(buttons, textvariable=other, width=digits + 3, font=(FONT, 12, "bold"),
                              style="Code.TEntry", validate="key", validatecommand=check)
            entry.pack(side="left")
            entry.bind("<FocusIn>", lambda e, v=var: v.set("other"))

        tk.Frame(c, bg=BORDER, height=1).pack(fill="x", pady=(4, 10))
        self._label(c, self.tr("msa_title"), size=11, bold=True).pack(anchor="w")
        self.msa_frame = tk.Frame(c, bg=CARD)
        self.msa_frame.pack(fill="x", pady=(4, 0))

    def _build_step_create(self):
        c = self._card(4, self.tr("step4_title"))
        for i in range(1, 5):
            self._label(c, f"{i}.  " + self.tr(f"step4_line{i}"), wraplength=760).pack(anchor="w", pady=1)

    def _fill_msa_section(self):
        for child in self.msa_frame.winfo_children():
            child.destroy()
        f = self.msa_frame
        info = self.sheet_info
        if info is None:
            self._label(f, self.tr("msa_need_file"), color=MUTED).pack(anchor="w")
            return
        if "pref" in info.found:
            self._label(f, self.tr("msa_with_column", column=self.column_vars["pref"].get().strip()),
                        wraplength=780).pack(anchor="w")
            ttk.Entry(f, textvariable=self.indicators_var).pack(fill="x", pady=(6, 2))
            self._label(f, self.tr("indicators_hint"), color=MUTED, wraplength=780).pack(anchor="w")
            return

        eligible = [code for code in info.countries if code in core.PREFERENCE_CODES]
        other = [code for code in info.countries if code not in core.PREFERENCE_CODES]
        self.country_vars = {code: self.country_vars.get(code) or tk.BooleanVar(value=False) for code in eligible}
        if not eligible:
            self._label(f, self.tr("msa_no_eu"), wraplength=780).pack(anchor="w")
        else:
            self._label(f, self.tr("msa_pick_countries"), wraplength=780).pack(anchor="w")
            boxes = tk.Frame(f, bg=CARD)
            boxes.pack(anchor="w", pady=(6, 0))
            for i, code in enumerate(eligible):
                pref = core.PREFERENCE_CODES[code]
                text = code if pref == "MSA" else f"{code} ({pref})"
                ttk.Checkbutton(boxes, text=text, variable=self.country_vars[code]).grid(
                    row=i // 8, column=i % 8, sticky="w", padx=(0, 18), pady=2)
        if other:
            self._label(f, self.tr("msa_outside_eu", countries=", ".join(other)), color=MUTED,
                        wraplength=780).pack(anchor="w", pady=(6, 0))

    # ------------------------------------------------------------------ state

    def _refresh(self):
        excel = self.excel_var.get().strip()
        if excel:
            self.file_name_label.configure(text=Path(excel).name, fg=TEXT)
        else:
            self.file_name_label.configure(text=self.tr("no_file"), fg=MUTED)
        info = self.sheet_info
        missing_required = []
        for key, required in core.FIELDS:
            status = self.column_status[key]
            name = self.column_vars[key].get().strip()
            if info is None:
                status.configure(text="")
            elif key in info.found:
                status.configure(text="✓ " + self.tr("found"), fg=OK)
            elif not required:
                status.configure(text="– " + self.tr("not_used"), fg=MUTED)
            else:
                status.configure(text="⚠ " + self.tr("will_be_empty"), fg=WARN_TEXT)
                missing_required.append(key)
            self.column_boxes[key].configure(values=info.headers if info else [])
        if info is None:
            self.file_status_label.configure(text=self.tr("step1_none") if not excel else "", fg=MUTED)
        elif missing_required:
            self.file_status_label.configure(text="⚠ " + self.tr("columns_missing_short"), fg=WARN_TEXT)
        else:
            self.file_status_label.configure(
                text="✓ " + self.tr("file_ready", sheet=info.sheet_name), fg=OK)
        self._fill_msa_section()

    def _inspect(self):
        self._inspect_job = None
        excel = self.excel_var.get().strip()
        self.sheet_info = None
        if excel:
            try:
                self.sheet_info = core.inspect_sheet(excel, self._column_names())
            except core.ConversionError as exc:
                messagebox.showerror(self.tr("error"), self.tr(exc.key, **exc.params))
                self.excel_var.set("")
        self._refresh()

    def _schedule_inspect(self):
        if self._inspect_job is not None:
            self.after_cancel(self._inspect_job)
        self._inspect_job = self.after(400, self._inspect)

    def _column_names(self):
        return {k: v.get().strip() for k, v in self.column_vars.items()}

    def _set_language(self, code):
        self.lang = code
        self._remember()
        self._build()

    def _remember(self):
        self.settings.update({
            "language": self.lang,
            "column_names": self._column_names(),
            "positive_indicators": self.indicators_var.get(),
            "constants": self.constants,
            "output_folder": self.output_var.get().strip(),
        })
        save_settings(self.settings)

    # ------------------------------------------------------------------ actions

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title=self.tr("choose_file"),
            filetypes=[(self.tr("excel_filter"), "*.xlsx *.xlsm *.xls"), (self.tr("all_files"), "*.*")],
        )
        if path:
            self.excel_var.set(path)
            self.country_vars = {}
            self._inspect()

    def _edit_constants(self):
        win = tk.Toplevel(self, bg=CARD, padx=18, pady=14)
        win.title(self.tr("constants_title"))
        win.transient(self)
        self._label(win, self.tr("constants_title"), size=13, bold=True).grid(row=0, column=0, columnspan=2, sticky="w")
        self._label(win, self.tr("constants_hint"), color=MUTED, wraplength=420).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))
        vars_ = {}
        for i, key in enumerate(core.DEFAULT_CONSTANTS, start=2):
            self._label(win, key).grid(row=i, column=0, sticky="w", pady=2)
            vars_[key] = tk.StringVar(value=self.constants.get(key, ""))
            ttk.Entry(win, textvariable=vars_[key], width=30).grid(row=i, column=1, sticky="ew", padx=(10, 0), pady=2)

        def reset():
            for k, v in core.DEFAULT_CONSTANTS.items():
                vars_[k].set(v)

        def save():
            self.constants = {k: v.get() for k, v in vars_.items()}
            self._remember()
            win.destroy()

        buttons = tk.Frame(win, bg=CARD)
        buttons.grid(row=len(vars_) + 2, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text=self.tr("reset"), style="Link.TButton", command=reset).pack(side="left", padx=4)
        ttk.Button(buttons, text=self.tr("cancel"), style="Link.TButton", command=win.destroy).pack(side="left", padx=4)
        ttk.Button(buttons, text=self.tr("save"), style="Big.TButton", command=save).pack(side="left", padx=4)
        win.grab_set()

    def _create(self):
        excel = self.excel_var.get().strip()
        if not excel:
            messagebox.showwarning(self.tr("missing_title"), self.tr("need_file"))
            return
        auto = self.precision_auto.get()
        if auto and not self.precision_var.get():
            precision = ""  # only needed if TARIK can't decide a code; convert() says so
        else:
            precision = self._answer(self.precision_var, self.precision_other, core.PRECISION_1_DIGITS,
                                     "need_precision")
            if precision is None:
                return
        extended = self._answer(self.extended_var, self.extended_other, core.EXTENDED_PROCEDURE_DIGITS,
                                "need_extended")
        if extended is None:
            return
        self._remember()

        options = core.Options(
            precision_1=precision,
            precision_auto=auto,
            extended_procedure=extended,
            column_names=self._column_names(),
            positive_indicators=[s.strip() for s in self.indicators_var.get().split(",") if s.strip()],
            msa_countries=[code for code, var in self.country_vars.items() if var.get()],
            constants=self.constants,
        )
        try:
            result = core.convert(excel, options)
        except core.ConversionError as exc:
            messagebox.showerror(self.tr("error"), self.tr(exc.key, **exc.params))
            return

        # Ask where to save every time; start in the folder used last time.
        folder = self.output_var.get().strip()
        if not folder or not Path(folder).is_dir():
            folder = str(Path(excel).parent)
        path = filedialog.asksaveasfilename(
            title=self.tr("save_title"),
            initialdir=folder,
            initialfile=core.output_filename(excel),
            defaultextension=".xml",
            filetypes=[(self.tr("xml_filter"), "*.xml"), (self.tr("all_files"), "*.*")],
        )
        if not path:
            return  # cancelled; the dialog itself asks before replacing a file
        out = Path(path)
        self.output_var.set(str(out.parent))
        self._remember()
        try:
            core.write_xml(result.xml, out)
        except OSError as exc:
            messagebox.showerror(self.tr("error"), self.tr("err_save", detail=str(exc)))
            return
        self._show_result(result, out)

    def _answer(self, var, other, digits, missing_key):
        """The chosen code, or None after telling the user what is missing."""
        value = var.get()
        if value == "other":
            value = other.get().strip()
            if not core.is_code(value, digits):
                messagebox.showwarning(self.tr("missing_title"), self.tr("need_other_code", digits=digits))
                return None
        if not value:
            messagebox.showwarning(self.tr("missing_title"), self.tr(missing_key))
            return None
        return value

    def _show_result(self, result, out: Path):
        win = tk.Toplevel(self, bg=CARD, padx=24, pady=20)
        win.title(self.tr("done"))
        win.transient(self)
        self.result_window = win
        head = tk.Frame(win, bg=CARD)
        head.pack(fill="x")
        tk.Label(head, text="✓", font=(FONT, 26, "bold"), fg="#FFFFFF", bg=OK, width=2).pack(side="left")
        titles = tk.Frame(head, bg=CARD)
        titles.pack(side="left", padx=12)
        self._label(titles, self.tr("result_title"), size=14, bold=True).pack(anchor="w")
        self._label(titles, out.name, color=MUTED).pack(anchor="w")

        stats = tk.Frame(win, bg=PRIMARY_LIGHT, padx=14, pady=10)
        stats.pack(fill="x", pady=(16, 0))
        rows = [
            (self.tr("stat_items"), str(len(result.items))),
            (self.tr("stat_msa"), str(result.msa_count)),
            (self.tr("stat_total"), core.format_number(result.total_amount, 2)),
            (self.tr("stat_gross"), core.format_number(result.total_gross, 3)),
            (self.tr("stat_net"), core.format_number(result.total_net, 3)),
        ]
        for i, (name, value) in enumerate(rows):
            self._label(stats, name, bg=PRIMARY_LIGHT).grid(row=i, column=0, sticky="w", pady=1)
            self._label(stats, value, size=11, bold=True, bg=PRIMARY_LIGHT).grid(
                row=i, column=1, sticky="e", padx=(40, 0), pady=1)
        self._label(win, self.tr("compare_totals"), color=MUTED, wraplength=460).pack(anchor="w", pady=(8, 0))

        if result.warnings:
            box = tk.Frame(win, bg=WARN_BG, padx=12, pady=10)
            box.pack(fill="both", expand=True, pady=(12, 0))
            self._label(box, "⚠ " + self.tr("warnings_title", count=len(result.warnings)), bold=True,
                        color=WARN_TEXT, bg=WARN_BG).pack(anchor="w")
            text = tk.Text(box, height=min(8, len(result.warnings) + 1), wrap="word", bg=WARN_BG, fg=WARN_TEXT,
                           relief="flat", font=(FONT, 10), width=60)
            text.insert("1.0", "\n".join("• " + self.tr(w.key, **w.params) for w in result.warnings))
            text.configure(state="disabled")
            text.pack(fill="both", expand=True, pady=(4, 0))

        buttons = tk.Frame(win, bg=CARD)
        buttons.pack(fill="x", pady=(16, 0))
        ttk.Button(buttons, text=self.tr("close"), style="Link.TButton", command=win.destroy).pack(side="right")
        ttk.Button(buttons, text="📂  " + self.tr("open_folder"), style="Big.TButton",
                   command=lambda: open_folder(out.parent)).pack(side="right", padx=8)
        win.grab_set()


def main():
    App().mainloop()
