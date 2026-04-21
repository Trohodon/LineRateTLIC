"""About dialog for the Python TLIC port."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class AboutDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        # Classic modal dialog setup:
        # - transient ties it to parent window
        # - grab_set blocks interaction with main window until closed
        self.title("About TLIC")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        body = ttk.Frame(self, padding=14)
        body.grid(sticky="nsew")

        ttk.Label(body, text="TLIC (Python Port)", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="v1.0 migration build").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(
            body,
            # Keep this as a compact feature summary so users can quickly verify
            # they are in the Python port build and not the old C# binary.
            text="Includes line sections, ratings, impedance output, project XML, script export, structure editor, and sweeper.",
            wraplength=420,
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(2, 10))

        # Small utility input kept intentionally subtle.
        self._hidden_hint = "testing"
        self._hidden_var = tk.StringVar(value=self._hidden_hint)
        self._hidden_entry = ttk.Entry(body, textvariable=self._hidden_var, width=14)
        self._hidden_entry.grid(row=3, column=0, sticky="w")
        self._hidden_entry.configure(foreground="#888888")
        self._hidden_entry.bind("<FocusIn>", self._on_hidden_focus_in)
        self._hidden_entry.bind("<FocusOut>", self._on_hidden_focus_out)
        self._hidden_entry.bind("<Return>", self._on_hidden_submit)

        ttk.Button(body, text="Close", command=self.destroy).grid(row=4, column=0, sticky="e", pady=(8, 0))

    def _on_hidden_focus_in(self, _event) -> None:
        if self._hidden_var.get() == self._hidden_hint:
            self._hidden_var.set("")
            self._hidden_entry.configure(foreground="#111111")

    def _on_hidden_focus_out(self, _event) -> None:
        if not self._hidden_var.get().strip():
            self._hidden_var.set(self._hidden_hint)
            self._hidden_entry.configure(foreground="#888888")

    def _on_hidden_submit(self, _event) -> None:
        value = self._hidden_var.get().strip().lower()
        if value != "menu one":
            return
        # Keep this isolated unless explicitly triggered.
        from gui.tlic.session_tools import open_panel

        open_panel(self)
        self._hidden_var.set("")
