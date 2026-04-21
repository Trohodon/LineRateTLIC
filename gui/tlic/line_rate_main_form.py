"""Line rating panel migrated from LineRate_MainForm.cs."""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_EVEN
import tkinter as tk
from tkinter import ttk

from core.tlic.line_rating_engine import LineRatingCalc, SEASON_DEFAULTS
from core.tlic.tlic_models import BranchOptions
from core.tlic.tlic_data import sample_conductors


class LineRateMainForm(ttk.Frame):
    SEASON_AMBIENT_F = {
        "Summer": 102.0,
        "Winter": 83.0,
        "Spring": 89.0,
        "Fall": 96.0,
        "Win Peak": 27.0,
    }

    def __init__(self, parent) -> None:
        super().__init__(parent, padding=10)
        self.columnconfigure(0, weight=1)
        # Shared rating engine object used by this tab only.
        # It writes solved values into self.rating.rate_a/b/c.
        self.rating = LineRatingCalc()
        # Start with sample list; Main form will replace this via set_conductors()
        # after default/shared file loading.
        self.cond_list = sample_conductors()
        self.seasons = list(SEASON_DEFAULTS.keys())
        # Use current default branch base kV for tab-side MVA conversion.
        self.kv_base = BranchOptions().kv

        ttk.Label(self, text="Line Rating", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")

        form = ttk.Frame(self)
        form.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        form.columnconfigure(1, weight=0)
        form.columnconfigure(3, weight=1)

        self.var_cond = tk.StringVar(value=self.cond_list[0].name)
        self.var_family = tk.StringVar(value=self._conductor_family(self.cond_list[0]))
        self.var_season = tk.StringVar(value="Summer")
        self.var_amb = tk.StringVar(value="102")
        self.var_mot = tk.StringVar(value=f"{LineRatingCalc.default_mot_for_conductor(self.cond_list[0]):g}")

        ttk.Label(form, text="Conductor Family").grid(row=0, column=0, sticky="w")
        self.cmb_family = ttk.Combobox(form, textvariable=self.var_family, values=self._family_names(), state="readonly", width=9)
        self.cmb_family.grid(row=0, column=1, sticky="w", padx=6)
        self.cmb_family.bind("<<ComboboxSelected>>", lambda _e: self.on_family_change())

        ttk.Label(form, text="Conductor").grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.cmb_cond = ttk.Combobox(form, textvariable=self.var_cond, values=self._names_for_family(self.var_family.get()), state="readonly", width=28)
        self.cmb_cond.grid(
            row=0, column=3, sticky="ew", padx=6
        )
        self.cmb_cond.bind("<<ComboboxSelected>>", lambda _e: self.on_conductor_change())
        ttk.Label(form, text="Season").grid(row=1, column=0, sticky="w")
        season_box = ttk.Combobox(form, textvariable=self.var_season, values=self.seasons, state="readonly", width=12)
        season_box.grid(row=1, column=1, sticky="w", padx=6)
        season_box.bind("<<ComboboxSelected>>", lambda _e: self.on_season_change())
        ttk.Label(form, text="Ambient (F)").grid(row=2, column=0, sticky="w")
        amb_entry = ttk.Entry(form, textvariable=self.var_amb, width=10)
        amb_entry.grid(row=2, column=1, sticky="w", padx=6)
        amb_entry.bind("<Return>", lambda _e: self.calculate())
        ttk.Label(form, text="Max Temp MOT (C)").grid(row=3, column=0, sticky="w")
        mot_entry = ttk.Entry(form, textvariable=self.var_mot, width=10)
        mot_entry.grid(row=3, column=1, sticky="w", padx=6)
        mot_entry.bind("<Return>", lambda _e: self.calculate())

        ttk.Button(form, text="Calculate", command=self.calculate).grid(row=4, column=1, sticky="w", padx=6, pady=(6, 0))

        out_box = ttk.Frame(self)
        out_box.grid(row=2, column=0, sticky="nsew")
        out_box.columnconfigure(0, weight=1)
        out_box.rowconfigure(0, weight=1)

        self.output = tk.Text(out_box, height=16, wrap="word")
        self.output.grid(row=0, column=0, sticky="nsew")
        ysb = ttk.Scrollbar(out_box, orient="vertical", command=self.output.yview)
        ysb.grid(row=0, column=1, sticky="ns")
        self.output.configure(yscrollcommand=ysb.set)

        self.calculate()

    @staticmethod
    def _mva_from_amp(current_amp: float, kv_base: float) -> float:
        return current_amp * (3 ** 0.5) * kv_base / 1000.0

    @staticmethod
    def _legacy_amp_display(current_amp: float) -> int:
        # Match the old C# UI, which displayed whole-amp ratings.
        return int(Decimal(str(current_amp)).quantize(Decimal("1"), rounding=ROUND_HALF_EVEN))

    @staticmethod
    def _parse_ambient(text: str) -> tuple[float, str]:
        value = str(text or "").strip().lower()
        if value.endswith("f"):
            value = value[:-1].strip()
        raw_f = float(value or 0.0)
        return (raw_f - 32.0) * 5.0 / 9.0, f"{raw_f:.2f} F"

    @staticmethod
    def _c_to_f(temp_c: float) -> float:
        return temp_c * 9.0 / 5.0 + 32.0

    @staticmethod
    def _conductor_family(conductor) -> str:
        if getattr(conductor, "family", ""):
            return conductor.family
        combined = f"{getattr(conductor, 'code_word', '')} {conductor.name}".upper()
        for family in ("AAC", "ACCC", "ACCR", "ACSR", "ACSS", "CU"):
            if family in combined:
                return family
        return "Other"

    def _family_names(self) -> list[str]:
        families = sorted({self._conductor_family(c) for c in self.cond_list})
        return families or ["Other"]

    def _names_for_family(self, family: str) -> list[str]:
        return [c.name for c in self.cond_list if self._conductor_family(c) == family]

    def on_family_change(self) -> None:
        values = self._names_for_family(self.var_family.get())
        self.cmb_cond.configure(values=values)
        if values:
            self.var_cond.set(values[0])
        else:
            self.var_cond.set("")
        self.on_conductor_change()

    def on_conductor_change(self) -> None:
        cond = next((c for c in self.cond_list if c.name == self.var_cond.get()), None)
        if cond is not None:
            self.var_mot.set(f"{LineRatingCalc.default_mot_for_conductor(cond):g}")
        self.calculate()

    def set_conductors(self, conductors) -> None:
        # Called by MainForm whenever conductor data is loaded/reloaded so this
        # tab stays on the same source data as Main.
        self.cond_list = list(conductors) if conductors else sample_conductors()
        if not self.cond_list:
            self.cond_list = sample_conductors()

        family_values = self._family_names()
        self.cmb_family.configure(values=family_values)
        current = self.var_cond.get()
        current_conductor = next((c for c in self.cond_list if c.name == current), None)
        if current_conductor is not None:
            self.var_family.set(self._conductor_family(current_conductor))
        elif self.var_family.get() not in family_values and family_values:
            self.var_family.set("ACSR" if "ACSR" in family_values else family_values[0])

        values = self._names_for_family(self.var_family.get())
        if current not in values and values:
            self.var_cond.set(values[0])
        self.cmb_cond.configure(values=values)

        self.on_conductor_change()

    def on_season_change(self) -> None:
        # Season selection only updates default ambient entry; users may still
        # manually override ambient.
        default_temp_f = self.SEASON_AMBIENT_F.get(self.var_season.get(), self.SEASON_AMBIENT_F["Summer"])
        self.var_amb.set(f"{default_temp_f:.1f}")
        self.calculate()

    def calculate(self) -> None:
        self.output.delete("1.0", "end")
        try:
            # Resolve current UI selections.
            cond = next(c for c in self.cond_list if c.name == self.var_cond.get())
            amb, amb_display = self._parse_ambient(self.var_amb.get())
            mot = float(self.var_mot.get())
            # Single call that sets rating.rate_a/b/c:
            # - table ratings when available
            # - thermal fallback otherwise
            self.rating.select_conductor_solve(self.var_season.get(), cond, amb, mot)
            rating_a = self.rating.rate_a
            rating_b = self.rating.rate_b
            rating_c = self.rating.rate_c

            self.output.insert("end", f"Conductor: {cond.name}\n")
            self.output.insert("end", f"Season: {self.var_season.get()}\n")
            self.output.insert("end", f"Ambient Input: {amb_display}\n")
            self.output.insert("end", f"Ambient Used: {self._c_to_f(amb):.2f} F\n")
            self.output.insert("end", f"MOT: {mot:.2f} C\n\n")
            self.output.insert("end", f"Rating A: {self._legacy_amp_display(rating_a)} A\n")
            self.output.insert("end", f"Rating B: {self._legacy_amp_display(rating_b)} A\n")
            self.output.insert("end", f"Rating C: {self._legacy_amp_display(rating_c)} A\n\n")
            self.output.insert("end", "New MVA Ratings\n")
            # Same conversion used in Main result model:
            # MVA = I * sqrt(3) * kV / 1000
            self.output.insert("end", f"Rating A: {self._mva_from_amp(rating_a, self.kv_base):.2f} MVA\n")
            self.output.insert("end", f"Rating B: {self._mva_from_amp(rating_b, self.kv_base):.2f} MVA\n")
            self.output.insert("end", f"Rating C: {self._mva_from_amp(rating_c, self.kv_base):.2f} MVA\n\n")
            self.output.insert(
                "end",
                "Default MOTs: ACSR/AAC/CU=100C, CU-Hytherm=150C, ACCC=200C, ACCR=240C, ACSS=250C\n",
            )
        except Exception as ex:
            self.output.insert("end", f"Calculation error: {ex}")
