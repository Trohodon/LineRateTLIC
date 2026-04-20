"""Variable sweeper migrated from TLineVariableSweeper.cs."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from core.tlic.line_rating_engine import ThermalModel


class VariableSweepException(Exception):
    # Raised when a sweep/plot combination is unsupported in this simplified
    # port implementation.
    pass


class VariableSweeperEngine:
    def __init__(self, model: ThermalModel) -> None:
        # Copy incoming model so each run can mutate sweep variables without
        # changing external caller state.
        self.model = ThermalModel(**model.__dict__)

    def sweep(self, to_sweep: str, to_plot: str, low: float, high: float, points: int = 100) -> tuple[list[float], list[float]]:
        # Generic N-point sweep loop.
        # x = swept variable values
        # y = requested plotted output for each x
        if points < 2:
            points = 2
        dx = (high - low) / points
        x = [low + i * dx for i in range(points)]
        y = [0.0 for _ in range(points)]

        # Track whether selected combination was recognized.
        found = False

        for i, xv in enumerate(x):
            # Each branch mutates one input, solves model, then records output.
            if to_sweep == "ConductorTemperature" and to_plot == "Current":
                self.model.max_temperature = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "Current" and to_plot == "ConductorTemperature":
                y[i] = self.model.temperature_at(xv, 0.1)
                found = True
            elif to_sweep == "Absorptivity" and to_plot == "Current":
                self.model.absorptivity = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "AmbientTemperature" and to_plot == "Current":
                self.model.ambient_temperature = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "Emissivity" and to_plot == "Current":
                self.model.emissivity = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "Diameter" and to_plot == "Current":
                self.model.diameter = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "WindVelocity" and to_plot == "Current":
                self.model.wind_velocity = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "LineWindAngle" and to_plot == "Current":
                self.model.line_wind_angle = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "DayOfYear" and to_plot == "Current":
                day = max(int(xv), 1)
                # Rough day->month mapping.
                self.model.month = min(12, max(1, ((day - 1) // 30) + 1))
                self.model.day_of_month = min(28, max(1, ((day - 1) % 30) + 1))
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True
            elif to_sweep == "HourOfDay" and to_plot == "Current":
                self.model.hour = xv
                self.model.solve_steady_state()
                y[i] = self.model.imax
                found = True

        if not found:
            raise VariableSweepException(f"Variable Combination: {to_sweep} and {to_plot} is not supported.")

        return x, y


class TLineVariableSweeper(ttk.Frame):
    def __init__(self, parent) -> None:
        super().__init__(parent, padding=10)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        ttk.Label(self, text="Variable Sweeper", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")

        controls = ttk.Frame(self)
        controls.grid(row=1, column=0, sticky="ew", pady=(8, 8))

        self.var_sweep = tk.StringVar(value="ConductorTemperature")
        self.var_plot = tk.StringVar(value="Current")
        self.var_low = tk.StringVar(value="80")
        self.var_high = tk.StringVar(value="250")

        choices = [
            "ConductorTemperature",
            "Current",
            "AmbientTemperature",
            "Emissivity",
            "Absorptivity",
            "Diameter",
            "WindVelocity",
            "LineWindAngle",
            "DayOfYear",
            "HourOfDay",
        ]

        ttk.Label(controls, text="Sweep").grid(row=0, column=0, sticky="w")
        # Keep choices readonly so branch logic can stay simple and explicit.
        ttk.Combobox(controls, textvariable=self.var_sweep, values=choices, width=22, state="readonly").grid(row=0, column=1, padx=6)
        ttk.Label(controls, text="Plot").grid(row=0, column=2, sticky="w")
        ttk.Combobox(controls, textvariable=self.var_plot, values=choices, width=22, state="readonly").grid(row=0, column=3, padx=6)
        ttk.Label(controls, text="Low").grid(row=0, column=4, sticky="w")
        ttk.Entry(controls, textvariable=self.var_low, width=8).grid(row=0, column=5, padx=6)
        ttk.Label(controls, text="High").grid(row=0, column=6, sticky="w")
        ttk.Entry(controls, textvariable=self.var_high, width=8).grid(row=0, column=7, padx=6)
        ttk.Button(controls, text="Run Sweep", command=self.run_sweep).grid(row=0, column=8, padx=8)

        self.canvas = tk.Canvas(self, background="#fcfcfc", highlightthickness=1, highlightbackground="#d0d0d0")
        self.canvas.grid(row=2, column=0, sticky="nsew")

        self.status = tk.StringVar(value="Sweep ready.")
        ttk.Label(self, textvariable=self.status, anchor="w").grid(row=3, column=0, sticky="ew", pady=(6, 0))

    def run_sweep(self) -> None:
        try:
            # Each run starts from a fresh ThermalModel to keep outputs
            # reproducible across repeated sweeps.
            low = float(self.var_low.get())
            high = float(self.var_high.get())
            model = ThermalModel()
            engine = VariableSweeperEngine(model)
            x, y = engine.sweep(self.var_sweep.get(), self.var_plot.get(), low, high)
            self._draw_xy(x, y)
            self.status.set(f"Sweep complete: {len(x)} points")
        except Exception as ex:
            self.status.set(f"Sweep error: {ex}")

    def _draw_xy(self, x: list[float], y: list[float]) -> None:
        self.canvas.delete("all")
        if not x or not y:
            return

        # Keep minimum size assumptions so first draw is legible.
        w = max(self.canvas.winfo_width(), 300)
        h = max(self.canvas.winfo_height(), 220)
        pad = 30

        xmin, xmax = min(x), max(x)
        ymin, ymax = min(y), max(y)
        if xmax - xmin < 1e-9:
            xmax = xmin + 1.0
        if ymax - ymin < 1e-9:
            ymax = ymin + 1.0

        def tx(v: float) -> float:
            return pad + (v - xmin) * (w - 2 * pad) / (xmax - xmin)

        def ty(v: float) -> float:
            return h - pad - (v - ymin) * (h - 2 * pad) / (ymax - ymin)

        # Axis lines
        self.canvas.create_line(pad, h - pad, w - pad, h - pad, fill="#444")
        self.canvas.create_line(pad, pad, pad, h - pad, fill="#444")

        # Polyline through all sampled points.
        pts: list[float] = []
        for xv, yv in zip(x, y):
            pts.extend([tx(xv), ty(yv)])
        self.canvas.create_line(*pts, fill="#c62828", width=2)

        self.canvas.create_text(pad + 6, pad + 6, text=f"Y [{ymin:.2f}, {ymax:.2f}]", anchor="nw", fill="#333")
        self.canvas.create_text(w - pad, h - pad + 12, text=f"X [{xmin:.2f}, {xmax:.2f}]", anchor="se", fill="#333")
