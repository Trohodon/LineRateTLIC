from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class Point:
    x: float = 0.0
    y: float = 0.0


@dataclass
class Structure:
    name: str
    a: list[Point] = field(default_factory=lambda: [Point(), Point(), Point()])
    g: list[Point] = field(default_factory=lambda: [Point(), Point()])

    @property
    def static_count(self) -> int:
        cnt = 0
        if self.g[0].y != 0.0:
            cnt += 1
        if self.g[1].y != 0.0:
            cnt += 1
        return cnt

    @property
    def es(self) -> float:
        pairs = [
            self._dist(self.a[0], self.a[1]),
            self._dist(self.a[1], self.a[2]),
            self._dist(self.a[0], self.a[2]),
        ]
        prod = 1.0
        for p in pairs:
            prod *= max(p, 1e-6)
        return prod ** (1.0 / 3.0)

    @staticmethod
    def _dist(p1: Point, p2: Point) -> float:
        dx = p1.x - p2.x
        dy = p1.y - p2.y
        return (dx * dx + dy * dy) ** 0.5


@dataclass
class Conductor:
    name: str
    aliases: list[str] = field(default_factory=list)
    family: str = ""
    code_word: str = ""
    ieee_conductor: Any | None = None
    # True when RATEA/B/C came from file/sample table.
    has_table_ratings: bool = False
    is_static_default: bool = False
    gmr_ft: float = 0.02
    radius_ft: float = 0.04
    r_ohm_per_mi: float = 0.08
    xl_ohm_per_mi: float = 0.35
    xc_mohm_mi: float = 0.20
    rate_a: float = 600.0
    rate_b: float = 700.0
    rate_c: float = 800.0
    od_in: float = 1.0
    r25_ohm_per_m: float = 0.00005
    r75_ohm_per_m: float = 0.000065
    heat_cap_ws_per_m_c: float = 500.0


@dataclass
class LineSection:
    cond_name: str
    static_name: str
    struct_name: str
    mileage: float
    is_custom_structure: bool = False
    mot: float = 125.0


@dataclass
class BranchOptions:
    bus1: int = 1
    bus2: int = 2
    ckt: str = "1"
    in_service: bool = True
    kv: float = 115.0
    mva_base: float = 100.0
    temp_c: float = 40.0
    rho: float = 100.0
    line_name: str = ""
    bus1_name: str = ""
    bus2_name: str = ""


@dataclass
class BranchResult:
    length_mi: float = 0.0
    r1_pu: float = 0.0
    x1_pu: float = 0.0
    b1_pu: float = 0.0
    r0_pu: float = 0.0
    x0_pu: float = 0.0
    b0_pu: float = 0.0
    z1_per_mile_r: float = 0.0
    z1_per_mile_x: float = 0.0
    y1_per_mile_b: float = 0.0
    z0_per_mile_r: float = 0.0
    z0_per_mile_x: float = 0.0
    y0_per_mile_b: float = 0.0
    raw_string: str = ""
    seq_string: str = ""
    current_rate_a: float = 0.0
    current_rate_b: float = 0.0
    current_rate_c: float = 0.0

    def mva_rating_a(self, kv: float) -> float:
        return self.current_rate_a * (3 ** 0.5) * kv / 1000.0

    def mva_rating_b(self, kv: float) -> float:
        return self.current_rate_b * (3 ** 0.5) * kv / 1000.0

    def mva_rating_c(self, kv: float) -> float:
        return self.current_rate_c * (3 ** 0.5) * kv / 1000.0


@dataclass
class ProjectData:
    options: BranchOptions = field(default_factory=BranchOptions)
    sections: list[LineSection] = field(default_factory=list)
    custom_structures: dict[str, Structure] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "options": asdict(self.options),
            "sections": [asdict(s) for s in self.sections],
            "custom_structures": {name: asdict(s) for name, s in self.custom_structures.items()},
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ProjectData":
        opts = BranchOptions(**data.get("options", {}))
        sections = [LineSection(**s) for s in data.get("sections", [])]
        custom_structures: dict[str, Structure] = {}
        for name, sdata in data.get("custom_structures", {}).items():
            custom_structures[name] = Structure(
                name=sdata.get("name", name),
                a=[Point(**p) for p in sdata.get("a", [{}, {}, {}])],
                g=[Point(**p) for p in sdata.get("g", [{}, {}])],
            )
        return cls(options=opts, sections=sections, custom_structures=custom_structures)
