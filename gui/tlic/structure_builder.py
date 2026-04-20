"""Structure builder UI migrated from StructureBuilder.cs."""

from __future__ import annotations

import copy
import tkinter as tk
from tkinter import ttk

from core.tlic.tlic_models import Point, Structure


class StructureEditorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, structure: Structure):
        super().__init__(parent)
        # This dialog edits one structure in memory and only commits on Save.
        # MainForm decides where/when the returned structure is persisted.
        self.title("Structure Builder")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # Keep both a baseline copy (for Reset) and a working copy (for edits).
        self.original = copy.deepcopy(structure)
        self.structure = copy.deepcopy(structure)
        self.result: Structure | None = None

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="Custom Name").grid(row=0, column=0, sticky="w")
        self.name_var = tk.StringVar(value=self.structure.name)
        ttk.Entry(top, textvariable=self.name_var, width=32).grid(row=0, column=1, sticky="w", padx=6)

        body = ttk.Frame(self, padding=(10, 0, 10, 10))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        form = ttk.LabelFrame(body, text="Coordinates")
        form.grid(row=0, column=0, sticky="nsw")
        self.vars: dict[str, tk.DoubleVar] = {}

        # Flat coordinate form to mirror old StructureBuilder workflow.
        labels = ["AX", "AY", "BX", "BY", "CX", "CY", "G1X", "G1Y", "G2X", "G2Y"]
        defaults = [
            self.structure.a[0].x,
            self.structure.a[0].y,
            self.structure.a[1].x,
            self.structure.a[1].y,
            self.structure.a[2].x,
            self.structure.a[2].y,
            self.structure.g[0].x,
            self.structure.g[0].y,
            self.structure.g[1].x,
            self.structure.g[1].y,
        ]

        for i, (lab, val) in enumerate(zip(labels, defaults)):
            ttk.Label(form, text=lab).grid(row=i, column=0, sticky="w", padx=6, pady=3)
            v = tk.DoubleVar(value=val)
            self.vars[lab] = v
            ent = ttk.Entry(form, textvariable=v, width=9)
            ent.grid(row=i, column=1, sticky="w", padx=6, pady=3)
            # Update live on typing and focus-out so the plot follows edits.
            ent.bind("<KeyRelease>", lambda _e: self._on_coord_change())
            ent.bind("<FocusOut>", lambda _e: self._on_coord_change())

        self.has_g1 = tk.BooleanVar(value=self.structure.g[0].y != 0.0)
        self.has_g2 = tk.BooleanVar(value=self.structure.g[1].y != 0.0)
        ttk.Checkbutton(form, text="Has G1", variable=self.has_g1, command=self._on_static_toggle).grid(
            row=10, column=0, sticky="w", padx=6, pady=3
        )
        ttk.Checkbutton(form, text="Has G2", variable=self.has_g2, command=self._on_static_toggle).grid(
            row=10, column=1, sticky="w", padx=6, pady=3
        )

        self.canvas = tk.Canvas(body, background="#fdfdfd", highlightthickness=1, highlightbackground="#d0d0d0")
        self.canvas.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        btns = ttk.Frame(self, padding=10)
        btns.grid(row=2, column=0, sticky="ew")
        ttk.Button(btns, text="Reset", command=self._reset).pack(side="left")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="Save", command=self._save).pack(side="right", padx=(0, 6))

        self._on_static_toggle()
        self._draw_structure()

    def _on_coord_change(self) -> None:
        try:
            # Move UI entry values into the working Structure object.
            self.structure.a[0] = Point(self.vars["AX"].get(), self.vars["AY"].get())
            self.structure.a[1] = Point(self.vars["BX"].get(), self.vars["BY"].get())
            self.structure.a[2] = Point(self.vars["CX"].get(), self.vars["CY"].get())
            self.structure.g[0] = Point(self.vars["G1X"].get(), self.vars["G1Y"].get())
            self.structure.g[1] = Point(self.vars["G2X"].get(), self.vars["G2Y"].get())
            self.structure.name = self.name_var.get().strip() or self.structure.name
            self._draw_structure()
        except tk.TclError:
            # Ignore transient partial-entry parse errors while user is typing.
            return

    def _on_static_toggle(self) -> None:
        # G2 is only valid when G1 exists. This keeps geometry coherent and
        # mirrors old UI behavior.
        g1_enabled = self.has_g1.get()
        g2_enabled = self.has_g2.get() and g1_enabled
        if not g1_enabled:
            self.has_g2.set(False)
            self.vars["G1X"].set(0.0)
            self.vars["G1Y"].set(0.0)
            self.vars["G2X"].set(0.0)
            self.vars["G2Y"].set(0.0)
        elif not g2_enabled:
            self.vars["G2X"].set(0.0)
            self.vars["G2Y"].set(0.0)
        self._on_coord_change()

    def _reset(self) -> None:
        # Full rollback to original structure passed into the dialog.
        self.structure = copy.deepcopy(self.original)
        self.name_var.set(self.structure.name)
        vals = {
            "AX": self.structure.a[0].x,
            "AY": self.structure.a[0].y,
            "BX": self.structure.a[1].x,
            "BY": self.structure.a[1].y,
            "CX": self.structure.a[2].x,
            "CY": self.structure.a[2].y,
            "G1X": self.structure.g[0].x,
            "G1Y": self.structure.g[0].y,
            "G2X": self.structure.g[1].x,
            "G2Y": self.structure.g[1].y,
        }
        for k, v in vals.items():
            self.vars[k].set(v)
        self.has_g1.set(self.structure.g[0].y != 0.0)
        self.has_g2.set(self.structure.g[1].y != 0.0)
        self._draw_structure()

    def _draw_structure(self) -> None:
        c = self.canvas
        c.delete("all")

        points = self.structure.a[:]
        if self.has_g1.get():
            points.append(self.structure.g[0])
        if self.has_g2.get():
            points.append(self.structure.g[1])

        # Maintain sane minimum drawing bounds for first paint.
        w = max(c.winfo_width(), 360)
        h = max(c.winfo_height(), 260)
        pad = 36

        xs = [p.x for p in points] + [0.0]
        ys = [p.y for p in points] + [0.0]
        raw_x_min, raw_x_max = min(xs), max(xs)
        raw_y_min, raw_y_max = min(ys), max(ys)

        x_span = max(raw_x_max - raw_x_min, 1.0)
        y_span = max(raw_y_max - raw_y_min, 1.0)
        x_pad = max(x_span * 0.15, 2.0)
        y_pad = max(y_span * 0.15, 2.0)

        x_min = raw_x_min - x_pad
        x_max = raw_x_max + x_pad
        y_min = min(0.0, raw_y_min - y_pad)
        y_max = raw_y_max + y_pad

        def tx(x: float) -> float:
            return pad + (x - x_min) * (w - 2 * pad) / (x_max - x_min)

        def ty(y: float) -> float:
            return h - pad - (y - y_min) * (h - 2 * pad) / (y_max - y_min)

        c.create_line(tx(0), ty(y_min), tx(0), ty(y_max), fill="#cccccc", width=2)
        c.create_line(tx(x_min), ty(0), tx(x_max), ty(0), fill="#cccccc", width=1)

        tick_count = 6
        for i in range(tick_count + 1):
            xv = x_min + (x_max - x_min) * i / tick_count
            xpix = tx(xv)
            c.create_line(xpix, ty(0) - 4, xpix, ty(0) + 4, fill="#999")
            c.create_text(xpix, ty(0) + 14, text=f"{xv:.1f}", fill="#666", anchor="n", font=("Segoe UI", 8))

            yv = y_min + (y_max - y_min) * i / tick_count
            ypix = ty(yv)
            c.create_line(tx(0) - 4, ypix, tx(0) + 4, ypix, fill="#999")
            c.create_text(tx(0) - 8, ypix, text=f"{yv:.1f}", fill="#666", anchor="e", font=("Segoe UI", 8))

        for label, p in zip(["A", "B", "C"], self.structure.a):
            x, y = tx(p.x), ty(p.y)
            c.create_oval(x - 5, y - 5, x + 5, y + 5, fill="#d32f2f", outline="")
            c.create_text(x + 10, y - 10, text=label, fill="#222", anchor="w")

        if self.has_g1.get():
            p = self.structure.g[0]
            x, y = tx(p.x), ty(p.y)
            c.create_oval(x - 4, y - 4, x + 4, y + 4, fill="#1565c0", outline="")
            c.create_text(x + 10, y - 10, text="G1", fill="#222", anchor="w")

        if self.has_g2.get():
            p = self.structure.g[1]
            x, y = tx(p.x), ty(p.y)
            c.create_oval(x - 4, y - 4, x + 4, y + 4, fill="#1565c0", outline="")
            c.create_text(x + 10, y - 10, text="G2", fill="#222", anchor="w")

        c.create_text(
            # Quick scale summary for manual coordinate tuning.
            tx(x_min) + 2,
            ty(y_min) + 2,
            text=f"X: {x_min:.1f}..{x_max:.1f}  Y: {y_min:.1f}..{y_max:.1f}",
            anchor="nw",
            fill="#666",
            font=("Segoe UI", 8),
        )

    def _save(self) -> None:
        # Commit edited structure to caller.
        self._on_coord_change()
        self.structure.name = self.name_var.get().strip() or self.structure.name
        self.result = copy.deepcopy(self.structure)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class StructureBuilder(ttk.Frame):
    def __init__(self, parent) -> None:
        super().__init__(parent, padding=10)
        # The dedicated tab is now informational only. Real editing is launched
        # from Main tab so edits stay tied to the selected structure context.
        self.columnconfigure(0, weight=1)
        ttk.Label(self, text="Structure Builder", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self,
            text="Use the Main tab 'Edit Structure' action to open the full editor for the currently selected structure.",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
