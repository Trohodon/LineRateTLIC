"""Main application window for the Python TLIC port."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .about_box import AboutDialog
from .structure_builder import StructureEditorDialog
from core.tlic.branch_engine import BranchEngine
from core.tlic.exporters import build_aux_script, build_python_script
from core.tlic.formatting import trunc_fixed
from core.tlic.line_rating_engine import LineRatingCalc, SEASON_DEFAULTS
from core.tlic.project_io import load_project_xml, save_project_xml
from core.tlic.tlic_data import by_name, load_conductors, load_static_conductors, load_structures, sample_conductors, sample_statics, sample_structures
from core.tlic.tlic_models import BranchOptions, Conductor, LineSection, ProjectData, Structure


class MainForm(ttk.Frame):
    SEASON_AMBIENT_F = {
        "Summer": 102.0,
        "Winter": 83.0,
        "Spring": 89.0,
        "Fall": 96.0,
        "Win Peak": 27.0,
    }

    def __init__(self, parent: tk.Misc, menu_parent: tk.Misc | None = None, build_menu: bool = True) -> None:
        super().__init__(parent)
        self.parent = parent
        self.menu_parent = menu_parent if menu_parent is not None else parent
        self.build_menu = build_menu

        self.kvs = [0.48, 0.6, 2.4, 4.0, 8.0, 12.0, 13.0, 13.8, 23.0, 33.0, 46.0, 69.0, 115.0, 230.0, 500.0]
        self.temps = [27.0, 83.0, 89.0, 96.0, 102.0]
        self.seasons = list(SEASON_DEFAULTS.keys())

        self.project = ProjectData(options=BranchOptions())
        self.phase_conds: list[Conductor] = sample_conductors()
        self.static_conds: list[Conductor] = sample_statics()
        self.structures: list[Structure] = sample_structures()

        self.external_cond_path = ""
        self.external_struct_path = ""

        self.rating_calc = LineRatingCalc()
        self.branch_engine = BranchEngine(self.rating_calc)
        self.last_result = None

        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()
        self._load_default_data()
        self._refresh_selectors()
        self.after_idle(self.on_structure_change)
        self.recalculate()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        if self.build_menu:
            self._build_menu()

        self.tab_main = ttk.Frame(self, padding=10)
        self.tab_main.grid(row=0, column=0, sticky="nsew", padx=10, pady=(8, 8))
        self._build_main_tab()

        status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        status.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

    def _build_menu(self) -> None:
        menu = tk.Menu(self.menu_parent)

        file_menu = tk.Menu(menu, tearoff=0)
        file_menu.add_command(label="Open Project...", command=self.on_open)
        file_menu.add_command(label="Save Project...", command=self.on_save)
        file_menu.add_separator()
        file_menu.add_command(label="Close Project", command=self.on_close_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.menu_parent.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        export_menu = tk.Menu(menu, tearoff=0)
        export_menu.add_command(label="Build Python Script...", command=self.on_export_python)
        export_menu.add_command(label="Build AUX Script...", command=self.on_export_aux)
        menu.add_cascade(label="Export", menu=export_menu)

        help_menu = tk.Menu(menu, tearoff=0)
        help_menu.add_command(label="About", command=self._open_about)
        menu.add_cascade(label="Help", menu=help_menu)

        self.menu_parent.configure(menu=menu)

    def _build_main_tab(self) -> None:
        self.tab_main.columnconfigure(1, weight=1)
        self.tab_main.rowconfigure(0, weight=1)

        left = ttk.Frame(self.tab_main)
        left.grid(row=0, column=0, sticky="nsw")

        right = ttk.Frame(self.tab_main)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.rowconfigure(3, weight=1)

        self._build_left_controls(left)
        self._build_right_views(right)

    def _build_left_controls(self, parent: ttk.Frame) -> None:
        row = 0

        files = ttk.LabelFrame(parent, text="Data Files", padding=8)
        files.grid(row=row, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)

        self.cond_file_var = tk.StringVar(value="")
        self.struct_file_var = tk.StringVar(value="")

        ttk.Label(files, text="Conductor").grid(row=0, column=0, sticky="w")
        ttk.Entry(files, textvariable=self.cond_file_var, width=36).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(files, text="Browse", command=self.on_browse_cond).grid(row=0, column=2)

        ttk.Label(files, text="Structure").grid(row=1, column=0, sticky="w")
        ttk.Entry(files, textvariable=self.struct_file_var, width=36).grid(row=1, column=1, sticky="ew", padx=4)
        ttk.Button(files, text="Browse", command=self.on_browse_struct).grid(row=1, column=2)

        row += 1
        selection = ttk.LabelFrame(parent, text="Selection", padding=8)
        selection.grid(row=row, column=0, sticky="ew", pady=(8, 0))
        selection.columnconfigure(1, weight=0)
        selection.columnconfigure(3, weight=1)

        self.cond_var = tk.StringVar()
        self.cond_family_var = tk.StringVar()
        self.static_var = tk.StringVar()
        self.struct_var = tk.StringVar()
        self.mot_var = tk.StringVar(value="100")
        self.temp_var = tk.StringVar(value="102.0")
        self.mileage_var = tk.StringVar(value="1.0")
        self.feet_var = tk.StringVar(value=str(5280.0))

        ttk.Label(selection, text="Conductor Family").grid(row=0, column=0, sticky="w")
        self.cmb_cond_family = ttk.Combobox(selection, textvariable=self.cond_family_var, state="readonly", width=9)
        self.cmb_cond_family.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        self.cmb_cond_family.bind("<<ComboboxSelected>>", lambda _e: self.on_cond_family_change())

        ttk.Label(selection, text="Conductor").grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.cmb_cond = ttk.Combobox(selection, textvariable=self.cond_var, state="readonly", width=24)
        self.cmb_cond.grid(row=0, column=3, sticky="ew", padx=4, pady=2)
        self.cmb_cond.bind("<<ComboboxSelected>>", lambda _e: self.on_conductor_change())

        ttk.Label(selection, text="Static").grid(row=1, column=0, sticky="w")
        self.cmb_static = ttk.Combobox(selection, textvariable=self.static_var, state="readonly", width=34)
        self.cmb_static.grid(row=1, column=1, columnspan=3, sticky="ew", padx=4, pady=2)
        self.cmb_static.bind("<<ComboboxSelected>>", lambda _e: self.on_static_change())

        ttk.Label(selection, text="Structure").grid(row=2, column=0, sticky="w")
        self.cmb_struct = ttk.Combobox(selection, textvariable=self.struct_var, state="readonly", width=34)
        self.cmb_struct.grid(row=2, column=1, columnspan=3, sticky="ew", padx=4, pady=2)
        self.cmb_struct.bind("<<ComboboxSelected>>", lambda _e: self.on_structure_change())

        ttk.Label(selection, text="MOT (C)").grid(row=3, column=0, sticky="w")
        mot = ttk.Entry(selection, textvariable=self.mot_var, width=10)
        mot.grid(row=3, column=1, sticky="w", padx=4, pady=2)
        mot.bind("<Return>", lambda _e: self.on_cond_or_season_change())

        ttk.Label(selection, text="Ambient (F)").grid(row=4, column=0, sticky="w")
        cmb_temp = ttk.Combobox(selection, textvariable=self.temp_var, values=[str(t) for t in self.temps], width=10)
        cmb_temp.grid(row=4, column=1, sticky="w", padx=4, pady=2)
        cmb_temp.bind("<<ComboboxSelected>>", lambda _e: self.on_cond_or_season_change())
        cmb_temp.bind("<Return>", lambda _e: self.on_cond_or_season_change())

        ttk.Label(selection, text="Mileage").grid(row=5, column=0, sticky="w")
        mil = ttk.Entry(selection, textvariable=self.mileage_var, width=12)
        mil.grid(row=5, column=1, sticky="w", padx=4, pady=2)
        mil.bind("<FocusOut>", lambda _e: self._sync_feet_from_miles())
        mil.bind("<Return>", lambda _e: self._sync_feet_from_miles())

        ttk.Label(selection, text="Feet").grid(row=6, column=0, sticky="w")
        feet = ttk.Entry(selection, textvariable=self.feet_var, width=12)
        feet.grid(row=6, column=1, sticky="w", padx=4, pady=2)
        feet.bind("<FocusOut>", lambda _e: self._sync_miles_from_feet())
        feet.bind("<Return>", lambda _e: self._sync_miles_from_feet())

        self.season_var = tk.StringVar(value="Summer")
        ttk.Label(selection, text="Season").grid(row=7, column=0, sticky="w", pady=(4, 0))
        cmb_season = ttk.Combobox(selection, textvariable=self.season_var, values=self.seasons, state="readonly", width=12)
        cmb_season.grid(row=7, column=1, sticky="w", padx=4, pady=(4, 0))
        cmb_season.bind("<<ComboboxSelected>>", lambda _e: self.on_season_change())

        row += 1
        branch = ttk.LabelFrame(parent, text="Branch Options", padding=8)
        branch.grid(row=row, column=0, sticky="ew", pady=(8, 0))
        branch.columnconfigure(1, weight=1)

        self.bus1_var = tk.StringVar(value="1")
        self.bus2_var = tk.StringVar(value="2")
        self.ckt_var = tk.StringVar(value="1")
        self.status_open_var = tk.BooleanVar(value=True)
        self.kv_var = tk.StringVar(value="115")
        self.mva_base_var = tk.StringVar(value="100")
        self.rho_var = tk.StringVar(value="100")
        self.line_name_var = tk.StringVar(value="")
        self.bus1_name_var = tk.StringVar(value="")
        self.bus2_name_var = tk.StringVar(value="")
        self.include_seq_var = tk.BooleanVar(value=True)

        labels = [
            ("Line Name", self.line_name_var),
            ("From Bus #", self.bus1_var),
            ("From Bus Name", self.bus1_name_var),
            ("To Bus #", self.bus2_var),
            ("To Bus Name", self.bus2_name_var),
            ("Circuit ID", self.ckt_var),
            ("kV Base", self.kv_var),
            ("MVA Base", self.mva_base_var),
            ("Ground Rho (ohm-m)", self.rho_var),
        ]

        for i, (lab, var) in enumerate(labels):
            ttk.Label(branch, text=lab).grid(row=i, column=0, sticky="w")
            if lab == "kV Base":
                ent = ttk.Combobox(branch, textvariable=var, values=[f"{kv:g}" for kv in self.kvs], state="readonly", width=14)
                ent.bind("<<ComboboxSelected>>", lambda _e: self.recalculate())
            else:
                ent = ttk.Entry(branch, textvariable=var, width=16)
                ent.bind("<FocusOut>", lambda _e: self.recalculate())
            ent.grid(row=i, column=1, sticky="ew", padx=4, pady=1)

        ttk.Checkbutton(branch, text="In Service", variable=self.status_open_var, command=self.recalculate).grid(
            row=9, column=0, sticky="w", pady=(4, 0)
        )
        ttk.Checkbutton(branch, text="Include Seq in Python Export", variable=self.include_seq_var).grid(
            row=9, column=1, sticky="w", pady=(4, 0)
        )

        row += 1
        actions = ttk.Frame(parent)
        actions.grid(row=row, column=0, sticky="ew", pady=(8, 0))

        ttk.Button(actions, text="Add Section", command=self.on_add_section).pack(side="left")
        ttk.Button(actions, text="Delete Selected", command=self.on_delete_selected).pack(side="left", padx=4)
        ttk.Button(actions, text="Clear Sections", command=self.on_clear_sections).pack(side="left", padx=4)
        ttk.Button(actions, text="Edit Structure", command=self.on_structure_edit).pack(side="left", padx=4)
        ttk.Button(actions, text="Recalculate", command=self.on_recalculate).pack(side="left", padx=4)
        ttk.Button(actions, text="Show Math", command=self.on_show_math).pack(side="left", padx=4)

    def _build_right_views(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(2, weight=2)

        struct_box = ttk.LabelFrame(parent, text="Structure Plot", padding=6)
        struct_box.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        struct_box.columnconfigure(0, weight=1)
        struct_box.rowconfigure(0, weight=1)

        self.struct_canvas = tk.Canvas(struct_box, background="#fefefe", height=230, highlightthickness=1, highlightbackground="#ddd")
        self.struct_canvas.grid(row=0, column=0, sticky="nsew")
        self.struct_canvas.bind("<Configure>", self._on_struct_canvas_configure)

        sections_box = ttk.LabelFrame(parent, text="Line Sections", padding=6)
        sections_box.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        sections_box.columnconfigure(0, weight=1)
        sections_box.rowconfigure(0, weight=1)

        cols = ("Struct", "Conductor", "Static", "MOT", "Mileage", "Custom")
        self.tree_sections = ttk.Treeview(sections_box, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            self.tree_sections.heading(c, text=c)
        self.tree_sections.column("Struct", width=100, anchor="w")
        self.tree_sections.column("Conductor", width=230, anchor="w")
        self.tree_sections.column("Static", width=180, anchor="w")
        self.tree_sections.column("MOT", width=70, anchor="center")
        self.tree_sections.column("Mileage", width=80, anchor="e")
        self.tree_sections.column("Custom", width=70, anchor="center")
        self.tree_sections.grid(row=0, column=0, sticky="nsew")

        ysb = ttk.Scrollbar(sections_box, orient="vertical", command=self.tree_sections.yview)
        ysb.grid(row=0, column=1, sticky="ns")
        self.tree_sections.configure(yscrollcommand=ysb.set)

        desc_box = ttk.LabelFrame(parent, text="Conductor/Static Details", padding=6)
        desc_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        desc_box.columnconfigure(0, weight=1)
        desc_box.columnconfigure(1, weight=1)
        desc_box.rowconfigure(0, weight=1)

        self.cond_desc = tk.Text(desc_box, height=8, wrap="word")
        self.cond_desc.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.static_desc = tk.Text(desc_box, height=8, wrap="word")
        self.static_desc.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        out_box = ttk.LabelFrame(parent, text="Impedance Calculation Output", padding=6)
        out_box.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        out_box.columnconfigure(0, weight=1)
        out_box.rowconfigure(0, weight=1)

        self.output = tk.Text(out_box, wrap="none")
        self.output.grid(row=0, column=0, sticky="nsew")
        y2 = ttk.Scrollbar(out_box, orient="vertical", command=self.output.yview)
        y2.grid(row=0, column=1, sticky="ns")
        x2 = ttk.Scrollbar(out_box, orient="horizontal", command=self.output.xview)
        x2.grid(row=1, column=0, sticky="ew")
        self.output.configure(yscrollcommand=y2.set, xscrollcommand=x2.set)

        ctx = tk.Menu(self.output, tearoff=0)
        ctx.add_command(label="Copy Selected", command=self._copy_output_selection)
        ctx.add_command(label="Select All", command=lambda: self.output.tag_add("sel", "1.0", "end"))
        self.output.bind("<Button-3>", lambda e: ctx.tk_popup(e.x_root, e.y_root))

    def _resource_path(self, filename: str) -> str:
        base_dir = os.path.dirname(os.path.dirname(__file__))
        project_dir = os.path.dirname(base_dir)
        project_resource = os.path.join(project_dir, "Resources", filename)
        return project_resource

    def _load_default_data(self) -> None:
        # Keep data documents centralized in the project Resources folder.
        cond_default = self._resource_path("ConData.xlsx")
        static_default = self._resource_path("StaticData.xlsx")
        struct_default = self._resource_path("StructureData.xlsx")
        if not os.path.exists(static_default):
            static_default = self._resource_path("StaticData.csv")
        if not os.path.exists(struct_default):
            struct_default = self._resource_path("structdata.txt")
        if not os.path.exists(struct_default):
            struct_default = ""

        self.cond_file_var.set(cond_default if os.path.exists(cond_default) else "(sample data)")
        self.struct_file_var.set(struct_default if os.path.exists(struct_default) else "(sample data)")

        self.phase_conds, embedded_statics = load_conductors(cond_default)
        self.static_conds = load_static_conductors(static_default) if os.path.exists(static_default) else embedded_statics
        self.structures = load_structures(struct_default)

    def _refresh_selectors(self) -> None:
        family_names = self._phase_family_names()
        static_names = [s.name for s in self.static_conds]

        self.cmb_cond_family.configure(values=family_names)
        current_cond = self.cond_var.get()
        if current_cond:
            selected = by_name(self.phase_conds, current_cond)
            if selected is not None:
                self.cond_family_var.set(self._conductor_family(selected))
        if self.cond_family_var.get() not in family_names and family_names:
            preferred = "ACSR" if "ACSR" in family_names else family_names[0]
            self.cond_family_var.set(preferred)

        cond_names = self._phase_cond_names_for_family(self.cond_family_var.get())
        self.cmb_cond.configure(values=cond_names)

        # Preserve legacy behavior for static selector: static wires first,
        # then phase conductors after a separator for unusual workflows.
        static_combo = static_names + ["----------"] + [c.name for c in self.phase_conds]

        self.cmb_static.configure(values=static_combo)

        struct_display = []
        self._struct_display_map: dict[str, str] = {}
        for s in self.structures + list(self.project.custom_structures.values()):
            label = f"{s.name:<11} ({s.es:.2f})"
            struct_display.append(label)
            self._struct_display_map[label] = s.name

        self.cmb_struct.configure(values=struct_display)

        if self.cond_var.get() not in cond_names and cond_names:
            self.cond_var.set(cond_names[0])
        elif not self.cond_var.get() and cond_names:
            self.cond_var.set(cond_names[0])
        if self.static_var.get() not in static_combo and static_names:
            self.static_var.set(static_names[0])
        elif not self.static_var.get() and static_combo:
            self.static_var.set(static_combo[0])
        if self.struct_var.get() not in struct_display and struct_display:
            self.struct_var.set(struct_display[0])
        elif not self.struct_var.get() and struct_display:
            self.struct_var.set(struct_display[0])

        self.on_structure_change()
        self.on_conductor_change()

    @staticmethod
    def _conductor_family(conductor: Conductor) -> str:
        if conductor.family:
            return conductor.family
        combined = f"{conductor.code_word} {conductor.name}".upper()
        for family in ("AAC", "ACCC", "ACCR", "ACSR", "ACSS", "CU"):
            if family in combined:
                return family
        return "Other"

    def _phase_family_names(self) -> list[str]:
        return sorted({self._conductor_family(c) for c in self.phase_conds})

    def _phase_cond_names_for_family(self, family: str) -> list[str]:
        return [c.name for c in self.phase_conds if self._conductor_family(c) == family]

    def on_cond_family_change(self) -> None:
        cond_names = self._phase_cond_names_for_family(self.cond_family_var.get())
        self.cmb_cond.configure(values=cond_names)
        if cond_names:
            self.cond_var.set(cond_names[0])
        else:
            self.cond_var.set("")
        self.on_conductor_change()

    def on_conductor_change(self) -> None:
        cond = by_name(self.phase_conds, self.cond_var.get())
        if cond is not None:
            self.mot_var.set(f"{LineRatingCalc.default_mot_for_conductor(cond):g}")
        self.on_cond_or_season_change()

    def on_browse_cond(self) -> None:
        path = filedialog.askopenfilename(
            title="Select conductor data file",
            filetypes=[("Conductor data", "*.xlsx *.csv *.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        self.external_cond_path = path
        self.cond_file_var.set(path)
        self.phase_conds, embedded_statics = load_conductors(path)
        if embedded_statics and embedded_statics != sample_statics():
            self.static_conds = embedded_statics
        self._refresh_selectors()
        self.recalculate()
        self.status_var.set(f"Loaded conductor data: {os.path.basename(path)}")

    def on_browse_struct(self) -> None:
        path = filedialog.askopenfilename(
            title="Select structure data file",
            filetypes=[("Structure data", "*.xlsx *.txt *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self.external_struct_path = path
        self.struct_file_var.set(path)
        self.structures = load_structures(path)
        self._refresh_selectors()
        self.recalculate()
        self.status_var.set(f"Loaded structure data: {os.path.basename(path)}")

    def _selected_structure_name(self) -> str:
        val = self.struct_var.get().strip()
        if val in self._struct_display_map:
            return self._struct_display_map[val]
        return val.split()[0] if val else ""

    def _selected_structure(self) -> Structure | None:
        name = self._selected_structure_name()
        if not name:
            return None
        if name in self.project.custom_structures:
            return self.project.custom_structures[name]
        return by_name(self.structures, name)

    def on_structure_change(self) -> None:
        self._draw_structure(self._selected_structure())

    def _on_struct_canvas_configure(self, _event) -> None:
        self._draw_structure(self._selected_structure())

    def _apply_season_default_temperature(self) -> None:
        default_temp_f = self.SEASON_AMBIENT_F.get(self.season_var.get(), self.SEASON_AMBIENT_F["Summer"])
        self.temp_var.set(f"{default_temp_f:.1f}")

    @staticmethod
    def _c_to_f(temp_c: float) -> float:
        return temp_c * 9.0 / 5.0 + 32.0

    @staticmethod
    def _parse_ambient_c(text: str) -> float:
        value = str(text or "").strip().lower()
        if value.endswith("f"):
            value = value[:-1].strip()
        return (float(value or 0.0) - 32.0) * 5.0 / 9.0

    def on_season_change(self) -> None:
        self._apply_season_default_temperature()
        self.on_cond_or_season_change()

    def on_cond_or_season_change(self) -> None:
        cond = by_name(self.phase_conds + self.static_conds, self.cond_var.get())
        if cond is None:
            return

        season = self.season_var.get()
        amb_c = self._parse_ambient_c(self.temp_var.get())
        mot_c = float(self.mot_var.get() or 125.0)
        self.rating_calc.select_conductor_solve(season, cond, amb_c, mot_c)

        self.cond_desc.delete("1.0", "end")
        self.cond_desc.insert("end", f"gmr:\t{cond.gmr_ft:.4f} ft\n")
        self.cond_desc.insert("end", f"rad:\t{cond.radius_ft:.4f} ft\n")
        self.cond_desc.insert("end", f"R:\t{cond.r_ohm_per_mi:.4f} ohm/mi\n")
        self.cond_desc.insert("end", f"XL:\t{cond.xl_ohm_per_mi:.4f} ohm/mi\n")
        self.cond_desc.insert("end", f"XC:\t{cond.xc_mohm_mi:.4f} Mohm-mi\n")
        self.cond_desc.insert("end", f"Rating A:\t{self.rating_calc.rate_a:.0f} A\n")
        self.cond_desc.insert("end", f"Rating B:\t{self.rating_calc.rate_b:.0f} A\n")
        self.cond_desc.insert("end", f"Rating C:\t{self.rating_calc.rate_c:.0f} A\n")

        self.on_static_change()
        self.recalculate()

    def on_static_change(self) -> None:
        st = by_name(self.phase_conds + self.static_conds, self.static_var.get())
        if st is None:
            return
        self.static_desc.delete("1.0", "end")
        self.static_desc.insert("end", f"gmr:\t{st.gmr_ft:.4f} ft\n")
        self.static_desc.insert("end", f"rad:\t{st.radius_ft:.4f} ft\n")
        self.static_desc.insert("end", f"R:\t{st.r_ohm_per_mi:.4f} ohm/mi\n")
        self.static_desc.insert("end", f"XL:\t{st.xl_ohm_per_mi:.4f} ohm/mi\n")
        self.static_desc.insert("end", f"XC:\t{st.xc_mohm_mi:.4f} Mohm-mi\n")
        self.static_desc.insert("end", f"Amp A:\t{st.rate_a:.0f} A\n")
        self.static_desc.insert("end", f"Amp B:\t{st.rate_b:.0f} A\n")
        self.static_desc.insert("end", f"Amp C:\t{st.rate_c:.0f} A\n")

    def _sync_feet_from_miles(self) -> None:
        try:
            self.feet_var.set(f"{float(self.mileage_var.get()) * 5280:.4f}")
        except Exception:
            pass

    def _sync_miles_from_feet(self) -> None:
        try:
            self.mileage_var.set(f"{float(self.feet_var.get()) / 5280:.6f}")
        except Exception:
            pass

    def _collect_options(self) -> BranchOptions:
        return BranchOptions(
            bus1=int(float(self.bus1_var.get() or 1)),
            bus2=int(float(self.bus2_var.get() or 2)),
            ckt=(self.ckt_var.get() or "1"),
            in_service=self.status_open_var.get(),
            kv=float(self.kv_var.get() or 230),
            mva_base=float(self.mva_base_var.get() or 100),
            temp_c=self._parse_ambient_c(self.temp_var.get() or 104),
            rho=float(self.rho_var.get() or 100),
            line_name=self.line_name_var.get(),
            bus1_name=self.bus1_name_var.get(),
            bus2_name=self.bus2_name_var.get(),
        )

    def on_add_section(self) -> None:
        try:
            sec = LineSection(
                cond_name=self.cond_var.get().strip(),
                static_name=self.static_var.get().strip(),
                struct_name=self._selected_structure_name(),
                mileage=float(self.mileage_var.get() or 1.0),
                is_custom_structure=self._selected_structure_name() in self.project.custom_structures,
                mot=float(self.mot_var.get() or 125),
            )
            self.project.sections.append(sec)
            self._refresh_sections_grid()
            children = self.tree_sections.get_children()
            if children:
                self.tree_sections.selection_set(children[-1])
                self.tree_sections.focus(children[-1])
            self.recalculate()
        except Exception as ex:
            messagebox.showerror("Add Section", f"Could not add section: {ex}")

    def on_delete_selected(self) -> None:
        selected = list(self.tree_sections.selection())
        if not selected:
            messagebox.showwarning("Delete Selected", "No selected sections to delete.")
            return
        if not messagebox.askokcancel("Delete Selected", "Delete selected line section(s)?"):
            return

        idxs = sorted((self.tree_sections.index(i) for i in selected), reverse=True)
        for idx in idxs:
            if 0 <= idx < len(self.project.sections):
                self.project.sections.pop(idx)

        self._refresh_sections_grid()
        self.recalculate()

    def on_clear_sections(self) -> None:
        if not messagebox.askyesno("Confirm", "Clear all recorded line sections?"):
            return
        self.project.sections.clear()
        self._refresh_sections_grid()
        self.recalculate()

    def on_structure_edit(self) -> None:
        current = self._selected_structure()
        if current is None:
            messagebox.showwarning("Structure Editor", "Select a structure first.")
            return
        dlg = StructureEditorDialog(self.winfo_toplevel(), current)
        self.wait_window(dlg)
        if dlg.result is None:
            return

        self.project.custom_structures[dlg.result.name] = dlg.result
        self.structures = [s for s in self.structures if s.name != dlg.result.name]
        self._refresh_selectors()
        # Select edited structure label.
        for label, name in self._struct_display_map.items():
            if name == dlg.result.name:
                self.struct_var.set(label)
                break
        self.on_structure_change()
        self.status_var.set(f"Custom structure saved: {dlg.result.name}")

    def _refresh_sections_grid(self) -> None:
        for iid in self.tree_sections.get_children():
            self.tree_sections.delete(iid)

        for sec in self.project.sections:
            self.tree_sections.insert(
                "",
                "end",
                values=(
                    sec.struct_name,
                    sec.cond_name,
                    sec.static_name,
                    f"{sec.mot:.1f}",
                    f"{sec.mileage:.3f}",
                    "True" if sec.is_custom_structure else "False",
                ),
            )

    def _current_section_inputs(self) -> LineSection:
        struct_name = self._selected_structure_name()
        return LineSection(
            cond_name=self.cond_var.get().strip(),
            static_name=self.static_var.get().strip(),
            struct_name=struct_name,
            mileage=float(self.mileage_var.get() or 1.0),
            is_custom_structure=struct_name in self.project.custom_structures,
            mot=float(self.mot_var.get() or 125),
        )

    def _selected_section_indexes(self) -> list[int]:
        indexes = sorted(self.tree_sections.index(iid) for iid in self.tree_sections.selection())
        return [idx for idx in indexes if 0 <= idx < len(self.project.sections)]

    def on_recalculate(self) -> None:
        try:
            target_indexes = self._selected_section_indexes()
            if not target_indexes and len(self.project.sections) == 1:
                target_indexes = [0]

            if target_indexes:
                updated = self._current_section_inputs()
                for idx in target_indexes:
                    self.project.sections[idx] = LineSection(
                        cond_name=updated.cond_name,
                        static_name=updated.static_name,
                        struct_name=updated.struct_name,
                        mileage=updated.mileage,
                        is_custom_structure=updated.is_custom_structure,
                        mot=updated.mot,
                    )
                self._refresh_sections_grid()
                children = self.tree_sections.get_children()
                reselection = [children[idx] for idx in target_indexes if 0 <= idx < len(children)]
                if reselection:
                    self.tree_sections.selection_set(reselection)
                    self.tree_sections.focus(reselection[0])

            self.recalculate()
        except Exception as ex:
            messagebox.showerror("Recalculate", f"Could not recalculate: {ex}")

    def recalculate(self) -> None:
        try:
            # Collect current branch-level inputs from UI (buses, kV, ambient,
            # MVA base, etc.) and run the branch engine over recorded sections.
            self.project.options = self._collect_options()
            structs = self.structures + list(self.project.custom_structures.values())
            self.last_result = self.branch_engine.calculate(
                self.project.options,
                self.project.sections,
                self.phase_conds,
                self.static_conds,
                structs,
                season=self.season_var.get(),
            )
            self._render_output()
            self.status_var.set("Calculation complete")
        except Exception as ex:
            self.status_var.set(f"Calculation failed: {ex}")

    def on_show_math(self) -> None:
        if not self.project.sections:
            messagebox.showinfo("Show Math", "Add at least one line section first.")
            return

        try:
            self.project.options = self._collect_options()
            structs = self.structures + list(self.project.custom_structures.values())
            report = self.branch_engine.build_math_report(
                self.project.options,
                self.project.sections,
                self.phase_conds,
                self.static_conds,
                structs,
                season=self.season_var.get(),
            )
        except Exception as ex:
            messagebox.showerror("Show Math", f"Could not build TLIC math view: {ex}")
            self.status_var.set("TLIC math view failed")
            return

        window = tk.Toplevel(self)
        window.title("TLIC Math")
        window.geometry("1040x760")
        window.minsize(820, 620)

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(frame)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(toolbar, text="Copy All", command=lambda: self._copy_text_to_clipboard(report)).pack(side="left")

        text = tk.Text(
            frame,
            wrap="none",
            font=("Consolas", 10),
            bg="#f6f1e8",
            fg="#18212b",
            insertbackground="#18212b",
            padx=16,
            pady=14,
        )
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)
        text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        text.grid(row=1, column=0, sticky="nsew")
        yscroll.grid(row=1, column=1, sticky="ns")
        xscroll.grid(row=2, column=0, sticky="ew")

        text.tag_configure("title", font=("Segoe UI", 18, "bold"), foreground="#102a43", spacing3=8)
        text.tag_configure("meta", font=("Segoe UI", 10, "italic"), foreground="#486581", spacing3=10)
        text.tag_configure("section", font=("Segoe UI", 13, "bold"), foreground="#7c2d12", spacing1=12, spacing3=6)
        text.tag_configure("subsection", font=("Segoe UI", 11, "bold"), foreground="#243b53", spacing1=8, spacing3=3)
        text.tag_configure("body", font=("Segoe UI", 10), foreground="#18212b", lmargin1=12, lmargin2=34, tabs=("30",))
        text.tag_configure("equation", font=("Segoe UI", 10, "bold"), foreground="#1f5f8b", lmargin1=12, lmargin2=34, tabs=("30",), spacing3=2)
        text.tag_configure("answer", font=("Segoe UI", 11, "bold"), foreground="#14532d", lmargin1=12, lmargin2=12, spacing3=2)
        text.tag_configure("highlight", font=("Segoe UI", 11, "bold"), foreground="#111827", background="#fde68a", lmargin1=8, lmargin2=8, spacing3=6)
        text.tag_configure("matrix", font=("Consolas", 10), foreground="#18212b", lmargin1=18, lmargin2=18, spacing3=1)

        self._populate_tlic_math_text(text, report)
        text.configure(state="disabled")
        self.status_var.set("Opened TLIC math view")

    def _copy_text_to_clipboard(self, text: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set("Copied TLIC math to clipboard")

    def _populate_tlic_math_text(self, text_widget: tk.Text, report: str) -> None:
        def insert(line: str = "", tag: str = "body") -> None:
            text_widget.insert("end", line + "\n", (tag,))

        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")

        lines = report.splitlines()
        i = 0
        while i < len(lines):
            line = lines[i]
            next_line = lines[i + 1] if i + 1 < len(lines) else ""

            if i == 0:
                insert(line, "title")
            elif i == 1:
                insert(line, "meta")
            elif next_line and set(next_line) == {"-"}:
                insert(line, "section")
                i += 1
            elif line.startswith("raw string:") or line.startswith("seq string:"):
                insert(line, "highlight")
            elif line.startswith(("Answer:", "Z1 =", "Y1 =", "Z0 =", "Y0 =", "R1/X1/B1", "R0/X0/B0", "Zero-sequence", "Total Z", "Total Y")):
                insert(line, "answer")
            elif line.endswith(":"):
                insert(line, "subsection")
            elif line in {
                "Conductor Data",
                "Structure Coordinates",
                "Positive Sequence Calculation",
                "Zero Sequence Series Calculation",
                "Zero Sequence Shunt Calculation",
                "Section Contribution",
            }:
                insert(line, "subsection")
            elif line.startswith(("Primitive ", "Kron-", "Sequence ", "Capacitance ")):
                insert(line, "subsection")
            elif line.startswith("  ["):
                insert(line, "matrix")
            elif line.startswith("  "):
                insert(line, "equation")
            elif line.startswith("- "):
                if any(token in line for token in ("ln(", "sqrt(", "inverse", "D_eq", "X1", "B1", "Zself", "Zmutual", "Zabc", "Zseq", "Pself", "Pmutual", "Pabc", "Cabc", "Yabc", "Yseq")):
                    insert(line, "equation")
                else:
                    insert(line, "body")
            elif "=" in line and any(token in line for token in ("ln(", "sqrt(", "inverse", "Zbase", "Ybase", "D_eq", "X1", "B1", "pu")):
                insert(line, "equation")
            elif not line.strip():
                insert("", "body")
            else:
                insert(line, "body")

            i += 1

    def _render_output(self) -> None:
        out = self.output
        out.delete("1.0", "end")

        if not self.project.sections:
            out.insert("end", "No Data\n")
            return

        res = self.last_result
        if res is None:
            out.insert("end", "No result\n")
            return

        lines = []
        lines.append("Impedance Calculation Output:\n")
        lines.append("Struct    Conductor                   Static                  MOT    Mileage")
        lines.append("--------------------------------------------------------------------------")
        for sec in self.project.sections:
            lines.append(
                f"{sec.struct_name:<9}{sec.cond_name:<28}{sec.static_name:<24}{sec.mot:<7.1f}{sec.mileage:.2f}"
            )

        lines.extend(
            [
                "",
                f"Total Length:\t\t{res.length_mi:.2f} mi",
                # Main rating printout path:
                # res.current_rate_* (amps) -> res.mva_rating_*(kV) -> shown MVA
                f"Rating A:\t\t{res.mva_rating_a(self.project.options.kv):.2f} MVA",
                f"Rating B:\t\t{res.mva_rating_b(self.project.options.kv):.2f} MVA",
                f"Rating C:\t\t{res.mva_rating_c(self.project.options.kv):.2f} MVA",
                f"Nominal Voltage:\t{self.project.options.kv:.2f} kV",
                f"MVA Base:\t\t{self.project.options.mva_base:.1f} MVA",
                "",
                "Per Unit Positive Sequence Impedances:",
                "----------------------------------------",
                f"R: {trunc_fixed(res.r1_pu)} p.u.",
                f"X: {trunc_fixed(res.x1_pu)} p.u.",
                f"B: {trunc_fixed(res.b1_pu)} p.u.",
                "",
                "Per Unit Zero Sequence Impedances:",
                "-----------------------------------",
                f"R0: {trunc_fixed(res.r0_pu)} p.u.",
                f"X0: {trunc_fixed(res.x0_pu)} p.u.",
                f"B0: {trunc_fixed(res.b0_pu)} p.u.",
                "",
                "Per Mile Impedances:",
                "-----------------------------------",
            ]
        )

        if len(self.project.sections) == 1:
            lines.extend(
                [
                    f"Z1: {trunc_fixed(res.z1_per_mile_r)} + j{trunc_fixed(res.z1_per_mile_x)} ohm/mi",
                    f"Y1: 0.000000 + j{trunc_fixed(res.y1_per_mile_b)} us/mi",
                    "",
                    f"Z0: {trunc_fixed(res.z0_per_mile_r)} + j{trunc_fixed(res.z0_per_mile_x)} ohm/mi",
                    f"Y0: 0.000000 + j{trunc_fixed(res.y0_per_mile_b)} us/mi",
                    "",
                ]
            )
        else:
            lines.append("(per mile impedances are not applicable to multi-section branches)")
            lines.append("")

        lines.extend(
            [
                "PSS/E Format:",
                "------------------------------",
                f"raw string:\n{res.raw_string} / {self.project.options.bus1_name} {self.project.options.bus2_name}",
                "",
                f"seq string:\n{res.seq_string}",
                "",
                "NOTE: sequence impedance uses a Python matrix engine:",
                "primitive overhead-line Z/Y, Carson-style earth return, grounded static-wire Kron reduction,",
                "and symmetrical-component conversion from the reduced phase matrix.",
            ]
        )

        out.insert("end", "\n".join(lines))

    def _draw_structure(self, structure: Structure | None) -> None:
        c = self.struct_canvas
        c.delete("all")
        if structure is None:
            return

        pts = structure.a[:] + [p for p in structure.g if p.y != 0.0]
        w = c.winfo_width()
        h = c.winfo_height()
        if w <= 1 or h <= 1:
            w = max(w, 350)
            h = max(h, 220)
        pad_x = max(24, min(40, int(w * 0.10)))
        pad_top = max(20, min(36, int(h * 0.10)))
        pad_bottom = max(34, min(50, int(h * 0.16)))

        xs = [p.x for p in pts] + [0.0]
        ys = [p.y for p in pts] + [0.0]
        raw_x_min, raw_x_max = min(xs), max(xs)
        raw_y_min, raw_y_max = min(ys), max(ys)

        x_span = max(raw_x_max - raw_x_min, 1.0)
        y_span = max(raw_y_max - raw_y_min, 1.0)
        x_pad = max(x_span * 0.18, 4.0)
        y_pad = max(y_span * 0.22, 4.0)

        x_min = raw_x_min - x_pad
        x_max = raw_x_max + x_pad
        y_min = max(0.0, raw_y_min - y_pad)
        y_max = raw_y_max + y_pad

        def tx(x: float) -> float:
            return pad_x + (x - x_min) * (w - 2 * pad_x) / (x_max - x_min)

        def ty(y: float) -> float:
            return h - pad_bottom - (y - y_min) * (h - pad_top - pad_bottom) / (y_max - y_min)

        # Draw axes.
        c.create_line(tx(0), ty(y_min), tx(0), ty(y_max), fill="#ddd", width=2)
        c.create_line(tx(x_min), ty(0), tx(x_max), ty(0), fill="#ddd", width=1)

        # Draw dynamic ticks/labels based on structure bounds.
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

        for lbl, p in zip(["A", "B", "C"], structure.a):
            x, y = tx(p.x), ty(p.y)
            c.create_oval(x - 5, y - 5, x + 5, y + 5, fill="#d32f2f", outline="")
            c.create_text(x + 8, y - 8, text=lbl, anchor="w", fill="#333")

        for i, p in enumerate(structure.g):
            if p.y == 0.0:
                continue
            x, y = tx(p.x), ty(p.y)
            c.create_oval(x - 4, y - 4, x + 4, y + 4, fill="#1565c0", outline="")
            c.create_text(x + 8, y - 8, text=f"G{i + 1}", anchor="w", fill="#333")
    def _copy_output_selection(self) -> None:
        try:
            txt = self.output.get("sel.first", "sel.last")
            self.clipboard_clear()
            self.clipboard_append(txt)
        except tk.TclError:
            return

    def on_save(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Project",
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self.project.options = self._collect_options()
            save_project_xml(path, self.project)
            self.status_var.set(f"Saved project: {os.path.basename(path)}")
        except Exception as ex:
            messagebox.showerror("Save", f"Could not save project: {ex}")

    def on_open(self) -> None:
        path = filedialog.askopenfilename(
            title="Open Project",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self.project = load_project_xml(path)
            self._apply_options_to_ui(self.project.options)
            self._refresh_selectors()
            self._refresh_sections_grid()
            self.recalculate()
            self.status_var.set(f"Opened project: {os.path.basename(path)}")
        except Exception as ex:
            messagebox.showerror("Open", f"Could not open project: {ex}")

    def _apply_options_to_ui(self, opts: BranchOptions) -> None:
        self.bus1_var.set(str(opts.bus1))
        self.bus2_var.set(str(opts.bus2))
        self.ckt_var.set(opts.ckt)
        self.status_open_var.set(opts.in_service)
        self.kv_var.set(str(opts.kv))
        self.mva_base_var.set(str(opts.mva_base))
        self.temp_var.set(f"{self._c_to_f(opts.temp_c):.1f}")
        self.rho_var.set(str(opts.rho))
        self.line_name_var.set(opts.line_name)
        self.bus1_name_var.set(opts.bus1_name)
        self.bus2_name_var.set(opts.bus2_name)

    def on_close_project(self) -> None:
        if not messagebox.askyesno("Close Project", "Reset all values to defaults and clear sections?"):
            return
        self.project = ProjectData(options=BranchOptions())
        self._apply_options_to_ui(self.project.options)
        self.mileage_var.set("1.0")
        self.mot_var.set("100")
        self._sync_feet_from_miles()
        self._refresh_sections_grid()
        self.recalculate()

    def on_export_python(self) -> None:
        if not self.project.sections or self.last_result is None:
            messagebox.showinfo("Export", "Cannot create output file. No line sections in list.")
            return

        path = filedialog.asksaveasfilename(
            title="Build Python File",
            defaultextension=".py",
            filetypes=[("Python Script", "*.py"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            content = build_python_script(self._collect_options(), self.last_result, self.include_seq_var.get())
            self._write_with_overwrite_or_append(path, content)
            self.status_var.set(f"Python script written: {os.path.basename(path)}")
        except Exception as ex:
            messagebox.showerror("Export", f"Could not write python script: {ex}")

    def on_export_aux(self) -> None:
        if not self.project.sections or self.last_result is None:
            messagebox.showinfo("Export", "Cannot create output file. No line sections in list.")
            return

        path = filedialog.asksaveasfilename(
            title="Build Aux File",
            defaultextension=".aux",
            filetypes=[("PowerWorld Auxiliary", "*.aux"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            content = build_aux_script(self._collect_options(), self.last_result)
            self._write_with_overwrite_or_append(path, content)
            self.status_var.set(f"AUX script written: {os.path.basename(path)}")
        except Exception as ex:
            messagebox.showerror("Export", f"Could not write aux script: {ex}")

    def _write_with_overwrite_or_append(self, path: str, content: str) -> None:
        mode = "w"
        if os.path.exists(path):
            choice = messagebox.askyesnocancel(
                "File exists",
                "File exists. Yes=overwrite, No=append, Cancel=cancel.",
            )
            if choice is None:
                return
            if choice is False:
                mode = "a"

        with open(path, mode, encoding="utf-8") as f:
            if mode == "a":
                f.write("\n\n")
            f.write(content)

    def _open_about(self) -> None:
        AboutDialog(self.winfo_toplevel())
