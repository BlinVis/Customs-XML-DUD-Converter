"""Tkinter desktop window for the converter (Albanian / English)."""

from __future__ import annotations

import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import core
from .i18n import LANGUAGES, t


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        self.lang = self.settings.get("language", "sq")
        if self.lang not in LANGUAGES:
            self.lang = "sq"

        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.precision_var = tk.StringVar(value="")  # no default: the user must answer
        self.extended_var = tk.StringVar(value="")
        saved_cols = self.settings.get("column_names", {})
        self.column_vars = {
            key: tk.StringVar(value=saved_cols.get(key, core.DEFAULT_COLUMN_NAMES[key]))
            for key, _ in core.FIELDS
        }
        self.indicators_var = tk.StringVar(
            value=self.settings.get("positive_indicators", ", ".join(core.DEFAULT_POSITIVE_INDICATORS))
        )
        self.constants = dict(core.DEFAULT_CONSTANTS)
        self.constants.update(self.settings.get("constants", {}))

        self.minsize(760, 720)
        self.geometry("860x860")
        self._build()
        self._apply_language()

    # ------------------------------------------------------------------ layout

    def _build(self):
        pad = {"padx": 8, "pady": 4}
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)

        top = ttk.Frame(root)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        self.title_label = ttk.Label(top, font=("Segoe UI", 14, "bold"))
        self.title_label.grid(row=0, column=0, sticky="w")
        self.lang_label = ttk.Label(top)
        self.lang_label.grid(row=0, column=1, sticky="e", padx=(0, 4))
        self.lang_box = ttk.Combobox(top, values=list(LANGUAGES.values()), state="readonly", width=9)
        self.lang_box.set(LANGUAGES[self.lang])
        self.lang_box.bind("<<ComboboxSelected>>", self._on_language)
        self.lang_box.grid(row=0, column=2, sticky="e")

        # Excel file
        self.file_frame = ttk.LabelFrame(root)
        self.file_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.file_frame.columnconfigure(0, weight=1)
        ttk.Entry(self.file_frame, textvariable=self.excel_var).grid(row=0, column=0, sticky="ew", **pad)
        self.browse_btn = ttk.Button(self.file_frame, command=self._browse_excel)
        self.browse_btn.grid(row=0, column=1, **pad)
        self.headers_btn = ttk.Button(self.file_frame, command=self._show_headers)
        self.headers_btn.grid(row=0, column=2, **pad)

        # Column names
        self.cols_frame = ttk.LabelFrame(root)
        self.cols_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.cols_frame.columnconfigure(1, weight=1)
        self.col_labels = {}
        for i, (key, _) in enumerate(core.FIELDS):
            lbl = ttk.Label(self.cols_frame)
            lbl.grid(row=i, column=0, sticky="w", padx=8, pady=2)
            ttk.Entry(self.cols_frame, textvariable=self.column_vars[key]).grid(
                row=i, column=1, sticky="ew", padx=8, pady=2
            )
            self.col_labels[key] = lbl

        # Questions
        self.q_frame = ttk.LabelFrame(root)
        self.q_frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.q_frame.columnconfigure(1, weight=1)
        self.precision_label = ttk.Label(self.q_frame)
        self.precision_label.grid(row=0, column=0, sticky="w", **pad)
        f = ttk.Frame(self.q_frame)
        f.grid(row=0, column=1, sticky="w")
        for v in core.PRECISION_1_CHOICES:
            ttk.Radiobutton(f, text=v, value=v, variable=self.precision_var).pack(side="left", padx=6)
        self.extended_label = ttk.Label(self.q_frame)
        self.extended_label.grid(row=1, column=0, sticky="w", **pad)
        f = ttk.Frame(self.q_frame)
        f.grid(row=1, column=1, sticky="w")
        for v in core.EXTENDED_PROCEDURE_CHOICES:
            ttk.Radiobutton(f, text=v, value=v, variable=self.extended_var).pack(side="left", padx=6)
        self.indicators_label = ttk.Label(self.q_frame)
        self.indicators_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 0))
        ttk.Entry(self.q_frame, textvariable=self.indicators_var).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=2
        )
        self.indicators_hint = ttk.Label(self.q_frame, foreground="#555")
        self.indicators_hint.grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        # Output folder
        self.out_frame = ttk.LabelFrame(root)
        self.out_frame.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.out_frame.columnconfigure(0, weight=1)
        ttk.Entry(self.out_frame, textvariable=self.output_var).grid(row=0, column=0, sticky="ew", **pad)
        self.out_btn = ttk.Button(self.out_frame, command=self._browse_output)
        self.out_btn.grid(row=0, column=1, **pad)
        self.out_hint = ttk.Label(self.out_frame, foreground="#555")
        self.out_hint.grid(row=1, column=0, sticky="w", padx=8)

        # Actions
        actions = ttk.Frame(root)
        actions.grid(row=5, column=0, sticky="ew", pady=8)
        self.constants_btn = ttk.Button(actions, command=self._edit_constants)
        self.constants_btn.pack(side="left")
        self.create_btn = ttk.Button(actions, command=self._create)
        self.create_btn.pack(side="right")

        # Log
        self.log_frame = ttk.LabelFrame(root)
        self.log_frame.grid(row=6, column=0, sticky="nsew")
        root.rowconfigure(6, weight=1)
        self.log = tk.Text(self.log_frame, height=10, wrap="word", state="disabled")
        scroll = ttk.Scrollbar(self.log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        scroll.pack(side="right", fill="y", pady=6)

    def _apply_language(self):
        L = self.lang
        self.title(t(L, "title"))
        self.title_label.configure(text=t(L, "title"))
        self.lang_label.configure(text=t(L, "language"))
        self.file_frame.configure(text=t(L, "excel_file"))
        self.browse_btn.configure(text=t(L, "browse"))
        self.headers_btn.configure(text=t(L, "show_headers"))
        self.cols_frame.configure(text=t(L, "columns_section"))
        for key, required in core.FIELDS:
            text = t(L, "field_" + key)
            if not required:
                text += " " + t(L, "optional")
            self.col_labels[key].configure(text=text)
        self.q_frame.configure(text=t(L, "questions_section"))
        self.precision_label.configure(text=t(L, "q_precision"))
        self.extended_label.configure(text=t(L, "q_extended"))
        self.indicators_label.configure(text=t(L, "q_indicators"))
        self.indicators_hint.configure(text=t(L, "indicators_hint"))
        self.out_frame.configure(text=t(L, "output_folder"))
        self.out_btn.configure(text=t(L, "browse"))
        self.out_hint.configure(text=t(L, "output_same"))
        self.constants_btn.configure(text=t(L, "constants"))
        self.create_btn.configure(text=t(L, "create"))
        self.log_frame.configure(text=t(L, "log"))

    def _on_language(self, _event=None):
        name = self.lang_box.get()
        self.lang = next(code for code, n in LANGUAGES.items() if n == name)
        self._apply_language()
        self._remember()

    # ------------------------------------------------------------------ actions

    def _write_log(self, text: str):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("1.0", text)
        self.log.configure(state="disabled")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[(t(self.lang, "excel_filter"), "*.xlsx *.xlsm *.xls"), (t(self.lang, "all_files"), "*.*")]
        )
        if path:
            self.excel_var.set(path)
            self._show_headers()

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _show_headers(self):
        path = self.excel_var.get().strip()
        if not path:
            messagebox.showwarning(t(self.lang, "error"), t(self.lang, "need_file"))
            return
        try:
            sheets = core.list_headers(path)
        except core.ConversionError as exc:
            messagebox.showerror(t(self.lang, "error"), t(self.lang, exc.key, **exc.params))
            return
        self._write_log("\n\n".join(
            t(self.lang, "headers_in", sheet=name, row=row, headers=" | ".join(headers))
            for name, row, headers in sheets
        ))

    def _edit_constants(self):
        win = tk.Toplevel(self)
        win.title(t(self.lang, "constants_title"))
        win.transient(self)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)
        vars_ = {}
        for i, key in enumerate(core.DEFAULT_CONSTANTS):
            ttk.Label(frame, text=key).grid(row=i, column=0, sticky="w", pady=2)
            vars_[key] = tk.StringVar(value=self.constants.get(key, ""))
            ttk.Entry(frame, textvariable=vars_[key], width=32).grid(row=i, column=1, sticky="ew", padx=6, pady=2)

        def reset():
            for k, v in core.DEFAULT_CONSTANTS.items():
                vars_[k].set(v)

        def save():
            self.constants = {k: v.get() for k, v in vars_.items()}
            self._remember()
            win.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=len(vars_), column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(buttons, text=t(self.lang, "reset"), command=reset).pack(side="left", padx=4)
        ttk.Button(buttons, text=t(self.lang, "cancel"), command=win.destroy).pack(side="left", padx=4)
        ttk.Button(buttons, text=t(self.lang, "save"), command=save).pack(side="left", padx=4)
        win.grab_set()

    def _remember(self):
        self.settings.update({
            "language": self.lang,
            "column_names": {k: v.get() for k, v in self.column_vars.items()},
            "positive_indicators": self.indicators_var.get(),
            "constants": self.constants,
        })
        save_settings(self.settings)

    def _create(self):
        L = self.lang
        excel = self.excel_var.get().strip()
        if not excel:
            messagebox.showwarning(t(L, "error"), t(L, "need_file"))
            return
        if not self.precision_var.get():
            messagebox.showwarning(t(L, "error"), t(L, "need_precision"))
            return
        if not self.extended_var.get():
            messagebox.showwarning(t(L, "error"), t(L, "need_extended"))
            return
        self._remember()

        options = core.Options(
            precision_1=self.precision_var.get(),
            extended_procedure=self.extended_var.get(),
            column_names={k: v.get().strip() for k, v in self.column_vars.items()},
            positive_indicators=[s.strip() for s in self.indicators_var.get().split(",") if s.strip()],
            constants=self.constants,
        )
        try:
            result = core.convert(excel, options)
        except core.ConversionError as exc:
            message = t(L, exc.key, **exc.params)
            self._write_log(message)
            messagebox.showerror(t(L, "error"), message)
            return

        folder = Path(self.output_var.get().strip() or Path(excel).parent)
        out = folder / core.output_filename(excel)
        if out.exists() and not messagebox.askyesno(t(L, "create"), t(L, "overwrite", name=out.name)):
            return
        try:
            core.write_xml(result.xml, out)
        except OSError as exc:
            messagebox.showerror(t(L, "error"), str(exc))
            return

        lines = [
            t(L, "saved", path=out),
            "",
            t(L, "summary", sheet=result.sheet_name, row=result.header_row, items=len(result.items),
              msa=result.msa_count, amount=core.format_number(result.total_amount, 2),
              gross=core.format_number(result.total_gross, 3), net=core.format_number(result.total_net, 3)),
        ]
        if result.indicator_values:
            values = [f"{v} → {t(L, 'positive' if pos else 'negative')}" for v, pos in result.indicator_values.items()]
            lines += ["", t(L, "indicators_found", values=values)]
        if result.warnings:
            lines += ["", t(L, "warnings")] + ["• " + t(L, w.key, **w.params) for w in result.warnings]
        self._write_log("\n".join(lines))
        messagebox.showinfo(t(L, "done"), t(L, "saved", path=out))


def main():
    App().mainloop()
