from __future__ import annotations

from dataclasses import dataclass
import math

from .formatting import trunc_fixed
from .line_rating_engine import LineRatingCalc
from .tlic_data import by_name
from .tlic_models import BranchOptions, BranchResult, Conductor, LineSection, Structure

FREQUENCY_HZ = 60.0
EPSILON_0_F_PER_M = 8.854e-12
METERS_PER_MILE = 1609.34
MICROSIEMENS_PER_SIEMENS = 1_000_000.0
CARSON_EARTH_RESISTANCE_OHM_PER_MILE_60HZ = 0.0953
CARSON_EARTH_DEPTH_FACTOR_FT = 2160.0
MIN_DISTANCE_FT = 1e-9


@dataclass
class SequencePerMile:
    z1: complex = 0j
    z0: complex = 0j
    y1: complex = 0j
    y0: complex = 0j


@dataclass
class BranchEngine:
    # Main-tab branch calculator.
    rating_calc: LineRatingCalc

    @staticmethod
    def _fmt(value: float, decimals: int = 6) -> str:
        return trunc_fixed(value, decimals)

    @classmethod
    def _fmt_complex(cls, value: complex) -> str:
        return f"{cls._fmt(value.real)} + j{cls._fmt(value.imag)}"

    @classmethod
    def _append_matrix(
        cls,
        lines: list[str],
        title: str,
        matrix: list[list[complex]],
        scale: float = 1.0,
        complex_values: bool = True,
    ) -> None:
        lines.append(title)
        for row in matrix:
            if complex_values:
                lines.append("  [" + ", ".join(cls._fmt_complex(value * scale) for value in row) + "]")
            else:
                lines.append("  [" + ", ".join(cls._fmt((value * scale).real) for value in row) + "]")
        lines.append("")

    @staticmethod
    def _phase_distance(p1, p2) -> float:
        dx = p1.x - p2.x
        dy = p1.y - p2.y
        return (dx * dx + dy * dy) ** 0.5

    @staticmethod
    def _mat_mul(a: list[list[complex]], b: list[list[complex]]) -> list[list[complex]]:
        rows = len(a)
        cols = len(b[0])
        inner = len(b)
        return [[sum(a[i][k] * b[k][j] for k in range(inner)) for j in range(cols)] for i in range(rows)]

    @staticmethod
    def _mat_sub(a: list[list[complex]], b: list[list[complex]]) -> list[list[complex]]:
        return [[a[i][j] - b[i][j] for j in range(len(a[0]))] for i in range(len(a))]

    @staticmethod
    def _mat_inv(a: list[list[complex]]) -> list[list[complex]]:
        n = len(a)
        aug = [[complex(a[i][j]) for j in range(n)] + [1.0 + 0j if i == j else 0j for j in range(n)] for i in range(n)]

        for col in range(n):
            pivot = max(range(col, n), key=lambda r: abs(aug[r][col]))
            if abs(aug[pivot][col]) < 1e-18:
                raise ValueError("matrix is singular")
            if pivot != col:
                aug[col], aug[pivot] = aug[pivot], aug[col]

            pv = aug[col][col]
            aug[col] = [v / pv for v in aug[col]]
            for row in range(n):
                if row == col:
                    continue
                factor = aug[row][col]
                if factor == 0:
                    continue
                aug[row] = [aug[row][j] - factor * aug[col][j] for j in range(2 * n)]

        return [row[n:] for row in aug]

    @classmethod
    def _kron_reduce(cls, matrix: list[list[complex]], keep_count: int) -> list[list[complex]]:
        total = len(matrix)
        eliminate_count = total - keep_count
        if eliminate_count <= 0:
            return [row[:] for row in matrix]

        pp = [row[:keep_count] for row in matrix[:keep_count]]
        pg = [row[keep_count:] for row in matrix[:keep_count]]
        gp = [row[:keep_count] for row in matrix[keep_count:]]
        gg = [row[keep_count:] for row in matrix[keep_count:]]
        return cls._mat_sub(pp, cls._mat_mul(cls._mat_mul(pg, cls._mat_inv(gg)), gp))

    @classmethod
    def _sequence_transform(cls, phase_matrix: list[list[complex]]) -> list[list[complex]]:
        a = complex(-0.5, math.sqrt(3.0) / 2.0)
        transform = [
            [1.0 + 0j, 1.0 + 0j, 1.0 + 0j],
            [1.0 + 0j, a * a, a],
            [1.0 + 0j, a, a * a],
        ]
        return cls._mat_mul(cls._mat_mul(cls._mat_inv(transform), phase_matrix), transform)

    @staticmethod
    def _validate_conductor(conductor: Conductor, label: str) -> None:
        if conductor.r_ohm_per_mi < 0.0:
            raise ValueError(f"{label} conductor resistance cannot be negative")
        if conductor.gmr_ft <= 0.0:
            raise ValueError(f"{label} conductor GMR must be greater than zero")
        if conductor.radius_ft <= 0.0:
            raise ValueError(f"{label} conductor radius must be greater than zero")

    @staticmethod
    def _distance_ft(p1, p2) -> float:
        dx = p1.x - p2.x
        dy = p1.y - p2.y
        return max((dx * dx + dy * dy) ** 0.5, MIN_DISTANCE_FT)

    @staticmethod
    def _image_distance_ft(p1, p2) -> float:
        dx = p1.x - p2.x
        dy = p1.y + p2.y
        return max((dx * dx + dy * dy) ** 0.5, MIN_DISTANCE_FT)

    @classmethod
    def _positive_sequence_per_mile(cls, conductor: Conductor, structure: Structure | None) -> tuple[float, float, float]:
        if structure is None:
            raise ValueError("structure geometry is required for positive-sequence impedance/admittance")

        a, b, c = structure.a
        d_ab = cls._phase_distance(a, b)
        d_bc = cls._phase_distance(b, c)
        d_ca = cls._phase_distance(c, a)
        if min(d_ab, d_bc, d_ca, conductor.gmr_ft, conductor.radius_ft) <= 0.0:
            raise ValueError(f"invalid positive-sequence geometry or conductor data for structure {structure.name}")

        d_eq = (d_ab * d_bc * d_ca) ** (1.0 / 3.0)
        x_ohm_per_mile = 0.12134 * math.log(d_eq / conductor.gmr_ft)
        capacitance_f_per_m = (2.0 * math.pi * EPSILON_0_F_PER_M) / math.log(d_eq / conductor.radius_ft)
        b_siemens_per_mile = 2.0 * math.pi * FREQUENCY_HZ * capacitance_f_per_m * METERS_PER_MILE
        return conductor.r_ohm_per_mi, x_ohm_per_mile, b_siemens_per_mile

    @classmethod
    def _conductors_for_structure(
        cls,
        phase_conductor: Conductor,
        static_conductor: Conductor,
        structure: Structure,
    ) -> tuple[list, list[Conductor]]:
        points = structure.a[:] + [p for p in structure.g if p.y != 0.0]
        conductors = [phase_conductor, phase_conductor, phase_conductor] + [static_conductor for _ in points[3:]]
        if len(points) < 3:
            raise ValueError(f"structure {structure.name} must have three phase coordinates")
        for i, p in enumerate(points[:3], start=1):
            if p.y <= 0.0:
                raise ValueError(f"structure {structure.name} phase {i} height must be greater than zero")
        for i, p in enumerate(points[3:], start=1):
            if p.y <= 0.0:
                raise ValueError(f"structure {structure.name} static wire {i} height must be greater than zero")
        return points, conductors

    @classmethod
    def _series_primitive_matrix(
        cls,
        phase_conductor: Conductor,
        static_conductor: Conductor,
        structure: Structure,
        rho_ohm_m: float,
    ) -> list[list[complex]]:
        cls._validate_conductor(phase_conductor, "phase")
        cls._validate_conductor(static_conductor, "static")
        points, conductors = cls._conductors_for_structure(phase_conductor, static_conductor, structure)

        rho = max(rho_ohm_m, 1e-6)
        earth_depth_ft = CARSON_EARTH_DEPTH_FACTOR_FT * math.sqrt(rho / FREQUENCY_HZ)
        earth_resistance = CARSON_EARTH_RESISTANCE_OHM_PER_MILE_60HZ * (FREQUENCY_HZ / 60.0)
        x_factor = 0.12134 * (FREQUENCY_HZ / 60.0)

        matrix: list[list[complex]] = []
        for i, cond_i in enumerate(conductors):
            row: list[complex] = []
            for j, cond_j in enumerate(conductors):
                if i == j:
                    row.append(complex(cond_i.r_ohm_per_mi + earth_resistance, x_factor * math.log(earth_depth_ft / cond_i.gmr_ft)))
                else:
                    dij = cls._distance_ft(points[i], points[j])
                    row.append(complex(earth_resistance, x_factor * math.log(earth_depth_ft / dij)))
            matrix.append(row)
        return matrix

    @classmethod
    def _shunt_potential_matrix(
        cls,
        phase_conductor: Conductor,
        static_conductor: Conductor,
        structure: Structure,
    ) -> list[list[complex]]:
        cls._validate_conductor(phase_conductor, "phase")
        cls._validate_conductor(static_conductor, "static")
        points, conductors = cls._conductors_for_structure(phase_conductor, static_conductor, structure)

        matrix: list[list[complex]] = []
        for i, cond_i in enumerate(conductors):
            row: list[complex] = []
            for j, _cond_j in enumerate(conductors):
                if i == j:
                    if points[i].y <= cond_i.radius_ft:
                        raise ValueError(f"conductor height must be greater than radius for structure {structure.name}")
                    value = math.log((2.0 * points[i].y) / cond_i.radius_ft) / (2.0 * math.pi * EPSILON_0_F_PER_M)
                else:
                    dij = cls._distance_ft(points[i], points[j])
                    dij_image = cls._image_distance_ft(points[i], points[j])
                    value = math.log(dij_image / dij) / (2.0 * math.pi * EPSILON_0_F_PER_M)
                row.append(complex(value, 0.0))
            matrix.append(row)
        return matrix

    @classmethod
    def _sequence_per_mile(
        cls,
        phase_conductor: Conductor,
        static_conductor: Conductor,
        structure: Structure | None,
        rho_ohm_m: float,
    ) -> SequencePerMile:
        if structure is None:
            raise ValueError("structure geometry is required for sequence impedance/admittance")

        z_primitive = cls._series_primitive_matrix(phase_conductor, static_conductor, structure, rho_ohm_m)
        z_phase = cls._kron_reduce(z_primitive, 3)
        z_sequence = cls._sequence_transform(z_phase)

        p_primitive = cls._shunt_potential_matrix(phase_conductor, static_conductor, structure)
        p_phase = cls._kron_reduce(p_primitive, 3)
        c_phase_f_per_m = cls._mat_inv(p_phase)
        omega = 2.0 * math.pi * FREQUENCY_HZ
        y_phase = [[1j * omega * c_phase_f_per_m[i][j] * METERS_PER_MILE for j in range(3)] for i in range(3)]
        y_sequence = cls._sequence_transform(y_phase)

        return SequencePerMile(
            z1=z_sequence[1][1],
            z0=z_sequence[0][0],
            y1=y_sequence[1][1],
            y0=y_sequence[0][0],
        )

    def build_math_report(
        self,
        options: BranchOptions,
        sections: list[LineSection],
        conductors: list[Conductor],
        statics: list[Conductor],
        structures: list[Structure],
        season: str,
    ) -> str:
        result = self.calculate(options, sections, conductors, statics, structures, season)
        if not sections:
            return "TLIC MATH\n\nNo line sections are in the branch."

        zbase = max((options.kv * options.kv) / max(options.mva_base, 0.001), 1e-6)
        ybase = 1.0 / zbase
        omega = 2.0 * math.pi * FREQUENCY_HZ
        earth_depth_ft = CARSON_EARTH_DEPTH_FACTOR_FT * math.sqrt(max(options.rho, 1e-6) / FREQUENCY_HZ)
        earth_resistance = CARSON_EARTH_RESISTANCE_OHM_PER_MILE_60HZ * (FREQUENCY_HZ / 60.0)

        given_lines: list[str] = []
        positive_lines: list[str] = []
        zero_lines: list[str] = []
        total_z1 = total_z0 = total_y1 = total_y0 = 0j
        total_miles = 0.0
        meta_parts: list[str] = []

        for idx, sec in enumerate(sections, start=1):
            ph = by_name(conductors + statics, sec.cond_name)
            st = by_name(statics + conductors, sec.static_name)
            structure = by_name(structures, sec.struct_name)
            if ph is None:
                given_lines.append(f"- Section {idx} skipped because conductor '{sec.cond_name}' was not found.")
                continue
            if st is None:
                st = ph
            if structure is None:
                given_lines.append(f"- Section {idx} skipped because structure '{sec.struct_name}' was not found.")
                continue
            if not meta_parts:
                meta_parts = [structure.name, ph.name, st.name, f"{self._fmt(sec.mileage)} mi"]

            miles = max(sec.mileage, 0.0)
            r1_mi, x1_mi, b1_mi = self._positive_sequence_per_mile(ph, structure)
            sequence = self._sequence_per_mile(ph, st, structure, options.rho)
            total_z1 += complex(r1_mi, x1_mi) * miles
            total_z0 += sequence.z0 * miles
            total_y1 += complex(0.0, b1_mi) * miles
            total_y0 += sequence.y0 * miles
            total_miles += miles

            a, b, c = structure.a
            d_ab = self._phase_distance(a, b)
            d_bc = self._phase_distance(b, c)
            d_ca = self._phase_distance(c, a)
            d_eq = (d_ab * d_bc * d_ca) ** (1.0 / 3.0)

            z_primitive = self._series_primitive_matrix(ph, st, structure, options.rho)
            z_phase = self._kron_reduce(z_primitive, 3)
            z_sequence = self._sequence_transform(z_phase)
            p_primitive = self._shunt_potential_matrix(ph, st, structure)
            p_phase = self._kron_reduce(p_primitive, 3)
            c_phase = self._mat_inv(p_phase)
            y_phase = [[1j * omega * c_phase[i][j] * METERS_PER_MILE for j in range(3)] for i in range(3)]
            y_sequence = self._sequence_transform(y_phase)

            given_lines.extend(
                [
                    f"Section {idx} Inputs:",
                    f"- Structure = {structure.name}",
                    f"- Phase conductor = {ph.name}",
                    f"- Static wire = {st.name}",
                    f"- Length = {self._fmt(miles)} mi",
                    f"- Phase R = {self._fmt(ph.r_ohm_per_mi)} ohm/mi",
                    f"- Phase GMR = {self._fmt(ph.gmr_ft)} ft",
                    f"- Phase radius = {self._fmt(ph.radius_ft)} ft",
                    f"- Static R = {self._fmt(st.r_ohm_per_mi)} ohm/mi",
                    f"- Static GMR = {self._fmt(st.gmr_ft)} ft",
                    f"- Static radius = {self._fmt(st.radius_ft)} ft",
                    "",
                    "Structure Coordinates:",
                    f"- A = ({self._fmt(a.x)}, {self._fmt(a.y)}) ft",
                    f"- B = ({self._fmt(b.x)}, {self._fmt(b.y)}) ft",
                    f"- C = ({self._fmt(c.x)}, {self._fmt(c.y)}) ft",
                ]
            )
            for gidx, point in enumerate([p for p in structure.g if p.y != 0.0], start=1):
                given_lines.append(f"- G{gidx} = ({self._fmt(point.x)}, {self._fmt(point.y)}) ft")
            given_lines.append("")

            positive_lines.extend(
                [
                    "",
                    f"Section {idx} Positive Sequence:",
                    "Phase spacing:",
                    f"  Dab = sqrt((xa - xb)^2 + (ya - yb)^2) = {self._fmt(d_ab)} ft",
                    f"  Dbc = sqrt((xb - xc)^2 + (yb - yc)^2) = {self._fmt(d_bc)} ft",
                    f"  Dca = sqrt((xc - xa)^2 + (yc - ya)^2) = {self._fmt(d_ca)} ft",
                    f"  D_eq = (Dab * Dbc * Dca)^(1/3) = {self._fmt(d_eq)} ft",
                    "",
                    "Series impedance:",
                    f"  R1_per_mile = conductor R = {self._fmt(r1_mi)} ohm/mi",
                    f"  X1_per_mile = 0.12134 * ln(D_eq / GMR)",
                    f"  X1_per_mile = 0.12134 * ln({self._fmt(d_eq)} / {self._fmt(ph.gmr_ft)}) = {self._fmt(x1_mi)} ohm/mi",
                    "",
                    "Shunt admittance:",
                    f"  C = (2 * pi * epsilon0) / ln(D_eq / radius)",
                    f"  B1_per_mile = 2 * pi * f * C * 1609.34 = {self._fmt(b1_mi * MICROSIEMENS_PER_SIEMENS)} us/mi",
                    "",
                    "Positive Sequence Answer:",
                    f"Answer: R1 = {self._fmt(result.r1_pu)} pu",
                    f"Answer: X1 = {self._fmt(result.x1_pu)} pu",
                    f"Answer: B1 = {self._fmt(result.b1_pu)} pu",
                    "",
                ]
            )

            zero_lines.extend(
                [
                    "",
                    f"Section {idx} Zero Sequence:",
                    "Carson-style series primitive:",
                    f"  De = 2160 * sqrt(rho / f) = {self._fmt(earth_depth_ft)} ft",
                    f"  Re = {self._fmt(earth_resistance)} ohm/mi",
                    f"  Zself = R + Re + j0.12134 * ln(De / GMR)",
                    f"  Zmutual = Re + j0.12134 * ln(De / Dij)",
                ]
            )
            self._append_matrix(zero_lines, "Primitive Z Matrix, ohm/mi, order [A, B, C, G1, G2...]", z_primitive)
            zero_lines.extend(
                [
                    "Static-wire Kron reduction:",
                    "  Zabc = Zpp - Zpg * inverse(Zgg) * Zgp",
                ]
            )
            self._append_matrix(zero_lines, "Kron-Reduced Phase Zabc, ohm/mi", z_phase)
            zero_lines.extend(
                [
                    "Symmetrical components:",
                    "  Zseq = inverse(A) * Zabc * A",
                ]
            )
            self._append_matrix(zero_lines, "Sequence Z Matrix, ohm/mi", z_sequence)
            zero_lines.extend(
                [
                    f"  Z0 = Zseq[0,0] = {self._fmt_complex(sequence.z0)} ohm/mi",
                    "",
                    "Image-distance shunt primitive:",
                    "  Pself = ln(2h / radius) / (2 * pi * epsilon0)",
                    "  Pmutual = ln(Dimage / Dij) / (2 * pi * epsilon0)",
                    "  Pabc = Ppp - Ppg * inverse(Pgg) * Pgp",
                    "  Cabc = inverse(Pabc)",
                    "  Yabc = j * 2 * pi * f * Cabc * 1609.34",
                    "  Yseq = inverse(A) * Yabc * A",
                ]
            )
            self._append_matrix(zero_lines, "Kron-Reduced Potential Pabc * 1e-9", p_phase, scale=1e-9, complex_values=False)
            self._append_matrix(zero_lines, "Capacitance Matrix Cabc = inverse(Pabc), nF/m", c_phase, scale=1e9, complex_values=False)
            self._append_matrix(zero_lines, "Sequence Y Matrix, us/mi", y_sequence, scale=MICROSIEMENS_PER_SIEMENS)
            zero_lines.extend(
                [
                    f"  Y0 = Yseq[0,0] = {self._fmt_complex(sequence.y0 * MICROSIEMENS_PER_SIEMENS)} us/mi",
                    "",
                    "Zero Sequence Answer:",
                    f"Answer: R0 = {self._fmt(result.r0_pu)} pu",
                    f"Answer: X0 = {self._fmt(result.x0_pu)} pu",
                    f"Answer: B0 = {self._fmt(result.b0_pu)} pu",
                    "",
                ]
            )

        lines: list[str] = [
            "TLIC IMPEDANCE AND ADMITTANCE MATH",
            " | ".join(meta_parts) if meta_parts else f"{len(sections)} section branch",
            "",
            "1. Given",
            "--------",
            f"From bus = {options.bus1}    To bus = {options.bus2}    Circuit = {options.ckt}",
            f"- Vbase = {self._fmt(options.kv)} kV",
            f"- Sbase = {self._fmt(options.mva_base)} MVA",
            f"- Zbase = Vbase^2 / Sbase = {self._fmt(zbase)} ohm",
            f"- Ybase = 1 / Zbase = {self._fmt(ybase)} S",
            f"- Frequency = {self._fmt(FREQUENCY_HZ)} Hz",
            f"- Ground rho = {self._fmt(options.rho)} ohm-m",
            f"- Carson equivalent earth depth De = {self._fmt(earth_depth_ft)} ft",
            f"- Carson earth resistance Re = {self._fmt(earth_resistance)} ohm/mi",
            "",
        ]
        lines.extend(given_lines)
        lines.extend(
            [
                "2. Positive Sequence",
                "--------------------",
            ]
        )
        lines.extend(positive_lines)
        lines.extend(
            [
                "3. Zero Sequence",
                "----------------",
            ]
        )
        lines.extend(zero_lines)
        lines.extend(
            [
                "4. Per-Mile Values And PSS/E Format",
                "-----------------------------------",
            ]
        )
        if len(sections) == 1:
            lines.extend(
                [
                    f"Z1 = {self._fmt_complex(complex(result.z1_per_mile_r, result.z1_per_mile_x))} ohm/mi",
                    f"Y1 = 0.000000 + j{self._fmt(result.y1_per_mile_b)} us/mi",
                    f"Z0 = {self._fmt_complex(complex(result.z0_per_mile_r, result.z0_per_mile_x))} ohm/mi",
                    f"Y0 = 0.000000 + j{self._fmt(result.y0_per_mile_b)} us/mi",
                ]
            )
        else:
            lines.append("Per-mile values are not shown for multi-section branches.")
        lines.extend(
            [
                "",
                "Branch Totals:",
                f"- Total length = {self._fmt(total_miles)} mi",
                f"- Total Z1 = {self._fmt_complex(total_z1)} ohm",
                f"- Total Z0 = {self._fmt_complex(total_z0)} ohm",
                f"- Total Y1 = {self._fmt_complex(total_y1 * MICROSIEMENS_PER_SIEMENS)} us",
                f"- Total Y0 = {self._fmt_complex(total_y0 * MICROSIEMENS_PER_SIEMENS)} us",
                "",
                "Per Unit Summary:",
                f"Answer: R1/X1/B1 = {self._fmt(result.r1_pu)} / {self._fmt(result.x1_pu)} / {self._fmt(result.b1_pu)} pu",
                f"Answer: R0/X0/B0 = {self._fmt(result.r0_pu)} / {self._fmt(result.x0_pu)} / {self._fmt(result.b0_pu)} pu",
                "",
                f"raw string: {result.raw_string}",
                f"seq string: {result.seq_string}",
            ]
        )

        return "\n".join(lines)

    def calculate(
        self,
        options: BranchOptions,
        sections: list[LineSection],
        conductors: list[Conductor],
        statics: list[Conductor],
        structures: list[Structure],
        season: str,
    ) -> BranchResult:
        result = BranchResult()
        if not sections:
            return result

        zbase = max((options.kv * options.kv) / max(options.mva_base, 0.001), 1e-6)
        bbase = 1.0 / zbase

        total_z1 = 0j
        total_z0 = 0j
        total_y1 = 0j
        total_y0 = 0j
        total_miles = 0.0

        # Branch rating is constrained by the weakest section.
        min_rate_a = min_rate_b = min_rate_c = 0.0

        for sec in sections:
            ph = by_name(conductors + statics, sec.cond_name)
            st = by_name(statics + conductors, sec.static_name)
            _str = by_name(structures, sec.struct_name)
            if ph is None:
                continue
            if st is None:
                st = ph

            miles = max(sec.mileage, 0.0)
            total_miles += miles

            r1_mi, x1_mi, b1_mi = self._positive_sequence_per_mile(ph, _str)
            sequence = self._sequence_per_mile(ph, st, _str, options.rho)

            total_z1 += complex(r1_mi, x1_mi) * miles
            total_z0 += sequence.z0 * miles
            total_y1 += complex(0.0, b1_mi) * miles
            total_y0 += sequence.y0 * miles

            self.rating_calc.select_conductor_solve(season, ph, options.temp_c, sec.mot)
            ra = self.rating_calc.rate_a
            rb = self.rating_calc.rate_b
            rc = self.rating_calc.rate_c

            if min_rate_a == 0.0 or ra < min_rate_a:
                min_rate_a = ra
            if min_rate_b == 0.0 or rb < min_rate_b:
                min_rate_b = rb
            if min_rate_c == 0.0 or rc < min_rate_c:
                min_rate_c = rc

        if total_miles <= 0:
            return result

        result.length_mi = total_miles

        result.r1_pu = total_z1.real / zbase
        result.x1_pu = total_z1.imag / zbase
        result.b1_pu = total_y1.imag / bbase

        result.r0_pu = total_z0.real / zbase
        result.x0_pu = total_z0.imag / zbase
        result.b0_pu = total_y0.imag / bbase

        if len(sections) == 1:
            z1_per_mile = total_z1 / total_miles
            z0_per_mile = total_z0 / total_miles
            y1_per_mile = total_y1 / total_miles
            y0_per_mile = total_y0 / total_miles
            result.z1_per_mile_r = z1_per_mile.real
            result.z1_per_mile_x = z1_per_mile.imag
            result.y1_per_mile_b = y1_per_mile.imag * MICROSIEMENS_PER_SIEMENS
            result.z0_per_mile_r = z0_per_mile.real
            result.z0_per_mile_x = z0_per_mile.imag
            result.y0_per_mile_b = y0_per_mile.imag * MICROSIEMENS_PER_SIEMENS

        result.current_rate_a = min_rate_a
        result.current_rate_b = min_rate_b
        result.current_rate_c = min_rate_c

        status = 1 if options.in_service else 0
        result.raw_string = (
            f"{options.bus1}, {options.bus2}, '{options.ckt}', {status}, "
            f"{trunc_fixed(result.r1_pu)}, {trunc_fixed(result.x1_pu)}, {trunc_fixed(result.b1_pu)}, "
            f"{result.mva_rating_a(options.kv):.2f}, {result.mva_rating_b(options.kv):.2f}, {result.mva_rating_c(options.kv):.2f}, {result.length_mi:.3f}"
        )
        result.seq_string = (
            f"{options.bus1}, {options.bus2}, '{options.ckt}', "
            f"{trunc_fixed(result.r0_pu)}, {trunc_fixed(result.x0_pu)}, {trunc_fixed(result.b0_pu)}"
        )

        return result
