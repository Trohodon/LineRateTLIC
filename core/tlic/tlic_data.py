from __future__ import annotations

import csv
import os
import re
import sys
from typing import Iterable

from core.line_rate.conductor_loader import load_conductor_database
from core.line_rate.xlsx_reader import list_sheet_names, read_sheet_table
from .tlic_models import Conductor, Point, Structure


def _num(value: str | None, default: float = 0.0) -> float:
    if value is None:
        return default
    txt = str(value).strip().replace("\ufeff", "")
    if txt == "":
        return default
    try:
        return float(txt)
    except ValueError:
        return default


def _canonical_header(value: str | None) -> str:
    txt = str(value or "").strip().replace("\ufeff", "").lower()
    txt = re.sub(r"\([^)]*\)", "", txt)
    return re.sub(r"[^a-z0-9]+", "", txt)


def _find_column_indexes(headers: list[str], *aliases: str) -> list[int]:
    wanted = {_canonical_header(alias) for alias in aliases if alias}
    return [idx for idx, header in enumerate(headers) if _canonical_header(header) in wanted]


def _get_cell(row: list[str], indexes: list[int], default: str = "") -> str:
    for idx in indexes:
        if 0 <= idx < len(row):
            value = row[idx].strip()
            if value != "":
                return value
    return default


def _get_last_cell(row: list[str], indexes: list[int], default: str = "") -> str:
    for idx in reversed(indexes):
        if 0 <= idx < len(row):
            value = row[idx].strip()
            if value != "":
                return value
    return default


def _sniff_delimiter(path: str) -> str:
    with open(path, "r", newline="", encoding="utf-8-sig", errors="ignore") as f:
        sample = f.read(4096)
    if "\t" in sample and sample.count("\t") > sample.count(","):
        return "\t"
    return ","


def _resource_path(*parts: str) -> str:
    base_dirs = []
    if getattr(sys, "frozen", False):
        base_dirs.append(os.path.dirname(sys.executable))
    base_dirs.append(getattr(sys, "_MEIPASS", os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))))

    for base in base_dirs:
        path = os.path.join(base, "Resources", *parts)
        if os.path.exists(path):
            return path
    return os.path.join(base_dirs[0], "Resources", *parts)


def _is_xlsx(path: str) -> bool:
    return os.path.splitext(path)[1].lower() == ".xlsx"


def _first_number(*values: float | None, default: float = 0.0) -> float:
    for value in values:
        if value is not None and value > 0.0:
            return value
    return default


def _average_resistance_ohm_per_mile(r25: float | None, r75: float | None, fallback: float | None = None) -> float:
    if r25 is not None and r75 is not None and r25 > 0.0 and r75 > 0.0:
        return (r25 + r75) / 2.0
    return _first_number(r75, r25, fallback, default=0.0)


def _display_name(code: str, ctype: str, size: str, name: str) -> str:
    if name:
        return name
    if size and size != "0":
        return f"{size} {ctype} ({code})".strip()
    return f"{code} ({ctype})".strip()


def _norm_key(value: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _code_base(value: str | None) -> str:
    text = str(value or "").strip()
    if "/" in text:
        text = text.split("/", 1)[0]
    return text


def _code_variants(value: str | None) -> list[str]:
    base = _code_base(value)
    variants = [base]
    upper = base.upper()
    for suffix in (" TW",):
        if upper.endswith(suffix):
            variants.append(base[: -len(suffix)])
    ordered: list[str] = []
    for variant in variants:
        if variant and variant not in ordered:
            ordered.append(variant)
    return ordered


def _lookup_keys(family: str, code: str, size: str = "", name: str = "") -> set[str]:
    fam = _norm_key(family)
    code_key = _norm_key(code)
    base_keys = [_norm_key(variant) for variant in _code_variants(code)]
    base_key = base_keys[0] if base_keys else ""
    size_key = _norm_key(size)
    name_key = _norm_key(name)

    keys = {key for key in (code_key, name_key, *base_keys) if key}
    if fam:
        for key in (code_key, name_key, *base_keys):
            if key:
                keys.add(f"{fam}|{key}")
        for key in (code_key, *base_keys):
            if size_key and key:
                keys.add(f"{fam}|{size_key}|{key}")
        if size_key:
            keys.add(f"{fam}|{size_key}")
    return keys


def _preferred_lookup_keys(family: str, code: str, size: str = "", name: str = "") -> list[str]:
    fam = _norm_key(family)
    code_key = _norm_key(code)
    base_keys = [_norm_key(variant) for variant in _code_variants(code)]
    base_key = base_keys[0] if base_keys else ""
    size_key = _norm_key(size)
    name_key = _norm_key(name)
    family_candidates = [
        f"{fam}|{size_key}|{code_key}" if fam and size_key and code_key else "",
        *[f"{fam}|{size_key}|{key}" for key in base_keys if fam and size_key and key],
        f"{fam}|{code_key}" if fam and code_key else "",
        *[f"{fam}|{key}" for key in base_keys if fam and key],
        f"{fam}|{name_key}" if fam and name_key else "",
    ]
    generic_candidates = [name_key, code_key, *base_keys]
    candidates = family_candidates if fam else generic_candidates
    ordered: list[str] = []
    for key in candidates:
        if key and key not in ordered:
            ordered.append(key)
    return ordered


def _row_electrical_values(row: list[str], indexes: dict[str, list[int]]) -> dict[str, float]:
    values = {
        "r": _num(_get_cell(row, indexes["r"]), -1.0),
        "xl": _num(_get_cell(row, indexes["xl"]), -1.0),
        "xc": _num(_get_cell(row, indexes["xc"]), -1.0),
        "gmr": _num(_get_cell(row, indexes["gmr"]), -1.0),
        "radius": _num(_get_cell(row, indexes["radius"]), -1.0),
    }
    return {key: value for key, value in values.items() if value >= 0.0}


def _merge_lookup(lookup: dict[str, dict[str, float]], keys: set[str], values: dict[str, float]) -> None:
    if not values:
        return
    for key in keys:
        if not key:
            continue
        merged = dict(lookup.get(key, {}))
        merged.update(values)
        lookup[key] = merged


def _rows_from_xlsx(path: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for sheet_name in list_sheet_names(path):
        table = read_sheet_table(path, sheet_name)
        if table:
            rows.extend([["" if value is None else str(value) for value in row] for row in table])
    return rows


def _electrical_lookup_from_xlsx(path: str) -> dict[str, dict[str, float]]:
    lookup: dict[str, dict[str, float]] = {}
    for sheet_name in list_sheet_names(path):
        table = read_sheet_table(path, sheet_name)
        if len(table) <= 1:
            continue
        headers = ["" if value is None else str(value) for value in table[0]]
        code_idxs = _find_column_indexes(headers, "code_name", "code")
        type_idxs = _find_column_indexes(headers, "type")
        size_idxs = _find_column_indexes(headers, "size")
        name_idxs = _find_column_indexes(headers, "name", "display_name")
        r_idxs = _find_column_indexes(headers, "r", "r_ohm_per_mi", "r(ohms/mi)", "resistance")
        xl_idxs = _find_column_indexes(headers, "xl", "x", "xl_ohm_per_mi", "xl_ohms_mi", "xl(ohms/mi)")
        xc_idxs = _find_column_indexes(
            headers,
            "xc",
            "c",
            "xc_ohms_mi",
            "xc_mohm_mi",
            "xc_mohms_mi",
            "xl_mohms_mi",
            "xc(ohms/mi)",
            "xc(ohms-mi)",
            "xc(mohm-mi)",
        )
        gmr_idxs = _find_column_indexes(headers, "gmr", "gmr_ft", "gmr(ft)")
        rad_idxs = _find_column_indexes(headers, "radius", "rad", "radius_ft", "radius(ft)")
        indexes = {"r": r_idxs, "xl": xl_idxs, "xc": xc_idxs, "gmr": gmr_idxs, "radius": rad_idxs}

        for raw in table[1:]:
            row = ["" if value is None else str(value).strip() for value in raw]
            keys = _lookup_keys(
                _get_cell(row, type_idxs),
                _get_cell(row, code_idxs),
                _get_cell(row, size_idxs),
                _get_last_cell(row, name_idxs),
            )
            _merge_lookup(lookup, keys, _row_electrical_values(row, indexes))
    return lookup


def _electrical_lookup_from_spaced_text(path: str) -> dict[str, dict[str, float]]:
    if not os.path.exists(path):
        return {}

    rows: list[list[str]] = []
    with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rows.append(re.split(r"\s{2,}", line))

    if len(rows) <= 1:
        return {}

    headers = rows[0]
    code_idxs = _find_column_indexes(headers, "code_name", "code")
    type_idxs = _find_column_indexes(headers, "type")
    size_idxs = _find_column_indexes(headers, "size", "name")
    name_idxs = _find_column_indexes(headers, "name", "display_name")
    r_idxs = _find_column_indexes(headers, "r", "r_ohm_per_mi", "r(ohms/mi)", "resistance")
    xl_idxs = _find_column_indexes(headers, "xl", "x", "xl_ohm_per_mi", "xl_ohms_mi", "xl(ohms/mi)")
    xc_idxs = _find_column_indexes(
        headers,
        "xc",
        "c",
        "xc_ohms_mi",
        "xc_mohm_mi",
        "xc_mohms_mi",
        "xl_mohms_mi",
        "xc(ohms/mi)",
        "xc(ohms-mi)",
        "xc(mohm-mi)",
    )
    gmr_idxs = _find_column_indexes(headers, "gmr", "gmr_ft", "gmr(ft)")
    rad_idxs = _find_column_indexes(headers, "radius", "rad", "radius_ft", "radius(ft)")
    indexes = {"r": r_idxs, "xl": xl_idxs, "xc": xc_idxs, "gmr": gmr_idxs, "radius": rad_idxs}

    lookup: dict[str, dict[str, float]] = {}
    for row in rows[1:]:
        keys = _lookup_keys(
            _get_cell(row, type_idxs),
            _get_cell(row, code_idxs),
            _get_cell(row, size_idxs),
            _get_last_cell(row, name_idxs),
        )
        _merge_lookup(lookup, keys, _row_electrical_values(row, indexes))
    return lookup


def _conductor_from_ieee(core_conductor, electrical: dict[str, float] | None = None) -> Conductor:
    electrical = electrical or {}
    code = core_conductor.code_word or ""
    family = core_conductor.family or ""
    size = "" if core_conductor.size_kcmil is None else f"{core_conductor.size_kcmil:g}"
    name = _display_name(code, family, size, core_conductor.name or "")
    radius_ft = electrical.get("radius", (core_conductor.od_in or 0.0) / 24.0)
    # If exact GMR is not supplied, use the solid-round approximation so the
    # field is visibly derived rather than silently copied from old conddata.
    gmr_ft = electrical.get("gmr", radius_ft * 0.7788 if radius_ft > 0.0 else 0.0)
    r_ohm_per_mi = electrical.get(
        "r",
        _average_resistance_ohm_per_mile(
            core_conductor.ac_res_25c_ohm_per_mile,
            core_conductor.ac_res_75c_ohm_per_mile,
            core_conductor.dc_res_20c_ohm_per_mile,
        ),
    )

    aliases = [value for value in {code, f"{size} {family} ({code})".strip()} if value and value != name]
    return Conductor(
        name=name,
        aliases=aliases,
        family=family,
        code_word=code,
        ieee_conductor=core_conductor,
        has_table_ratings=False,
        is_static_default=False,
        gmr_ft=gmr_ft,
        radius_ft=radius_ft,
        r_ohm_per_mi=r_ohm_per_mi,
        xl_ohm_per_mi=electrical.get("xl", 0.0),
        xc_mohm_mi=electrical.get("xc", 0.0),
        rate_a=0.0,
        rate_b=0.0,
        rate_c=0.0,
        od_in=(core_conductor.od_in or 0.0) * 25.4,
        r25_ohm_per_m=(core_conductor.ac_res_25c_ohm_per_mile or 0.0) / 1609.344,
        r75_ohm_per_m=(core_conductor.ac_res_75c_ohm_per_mile or 0.0) / 1609.344,
    )


def load_conductors_from_condata(path: str) -> list[Conductor]:
    database = load_conductor_database(path)
    electrical_lookup = _electrical_lookup_from_xlsx(path)
    supplemental_paths = [
        _resource_path("ConductorData.xlsx"),
    ]
    for supplemental_path in supplemental_paths:
        if not os.path.exists(supplemental_path) or os.path.abspath(supplemental_path) == os.path.abspath(path):
            continue
        if _is_xlsx(supplemental_path):
            supplemental_lookup = _electrical_lookup_from_xlsx(supplemental_path)
        else:
            supplemental_lookup = _electrical_lookup_from_spaced_text(supplemental_path)
        for key, values in supplemental_lookup.items():
            merged = dict(values)
            merged.update(electrical_lookup.get(key, {}))
            electrical_lookup[key] = merged
    conductors: list[Conductor] = []
    for family in database.get_families():
        for core_conductor in database.get_conductors(family):
            size = "" if core_conductor.size_kcmil is None else f"{core_conductor.size_kcmil:g}"
            electrical: dict[str, float] = {}
            for key in _preferred_lookup_keys(
                core_conductor.family or family,
                core_conductor.code_word or "",
                size,
                core_conductor.name or "",
            ):
                values = electrical_lookup.get(key)
                if values:
                    electrical.update(values)
                    break
            conductors.append(_conductor_from_ieee(core_conductor, electrical))
    return conductors


def _load_static_rows(rows: list[list[str]]) -> list[Conductor]:
    if len(rows) <= 1:
        return []
    headers = rows[0]
    display_name_idxs = _find_column_indexes(headers, "display_name", "name", "conductor", "condname", "description")
    code_idxs = _find_column_indexes(headers, "code_name", "code")
    type_idxs = _find_column_indexes(headers, "type")
    r_idxs = _find_column_indexes(headers, "r", "r_ohm_per_mi", "resistance", "r60", "r(ohms/mi)")
    xl_idxs = _find_column_indexes(headers, "xl", "x", "xl_ohm_per_mi", "xl_ohms_mi", "xl(ohms/mi)")
    xc_idxs = _find_column_indexes(
        headers,
        "xc",
        "c",
        "xc_ohms_mi",
        "xc_mohm_mi",
        "xc_mohms_mi",
        "xl_mohms_mi",
        "xc(ohms/mi)",
        "xc(ohms-mi)",
        "xc(mohm-mi)",
    )
    gmr_idxs = _find_column_indexes(headers, "gmr", "gmr_ft", "gmr(ft)")
    rad_idxs = _find_column_indexes(headers, "radius", "rad", "radius_ft", "radius(ft)")
    rate_a_idxs = _find_column_indexes(headers, "ratea", "rate_a", "ampa", "ratea(a)")
    rate_b_idxs = _find_column_indexes(headers, "rateb", "rate_b", "ampb", "rateb(a)")
    rate_c_idxs = _find_column_indexes(headers, "ratec", "rate_c", "ampc", "ratec(a)")

    statics: list[Conductor] = []
    for raw in rows[1:]:
        row = ["" if value is None else str(value).strip() for value in raw]
        name = _get_last_cell(row, display_name_idxs)
        code = _get_cell(row, code_idxs)
        ctype = _get_cell(row, type_idxs)
        if not name:
            name = _display_name(code, ctype, "", "")
        if not name:
            continue
        statics.append(
            Conductor(
                name=name,
                aliases=[value for value in {code, _get_cell(row, display_name_idxs)} if value and value != name],
                family=ctype,
                code_word=code,
                is_static_default=True,
                has_table_ratings=True,
                gmr_ft=_num(_get_cell(row, gmr_idxs), 0.0),
                radius_ft=_num(_get_cell(row, rad_idxs), 0.0),
                r_ohm_per_mi=_num(_get_cell(row, r_idxs), 0.0),
                xl_ohm_per_mi=_num(_get_cell(row, xl_idxs), 0.0),
                xc_mohm_mi=_num(_get_cell(row, xc_idxs), 0.0),
                rate_a=_num(_get_cell(row, rate_a_idxs), 0.0),
                rate_b=_num(_get_cell(row, rate_b_idxs), 0.0),
                rate_c=_num(_get_cell(row, rate_c_idxs), 0.0),
            )
        )
    return statics


def load_static_conductors(path: str) -> list[Conductor]:
    if not os.path.exists(path):
        return sample_statics()
    if _is_xlsx(path):
        statics = _load_static_rows(_rows_from_xlsx(path))
    else:
        delim = _sniff_delimiter(path)
        with open(path, "r", newline="", encoding="utf-8-sig", errors="ignore") as f:
            statics = _load_static_rows(list(csv.reader(f, delimiter=delim)))
    return statics or sample_statics()


def load_thermal_conductor_lookup(path: str | None = None) -> dict[str, dict[str, float]]:
    if path is None:
        return {}
    thermal_path = path
    if not os.path.exists(thermal_path):
        return {}

    lookup: dict[str, dict[str, float]] = {}
    with open(thermal_path, "r", newline="", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) <= 1:
        return lookup

    headers = rows[0]
    code_idx = _find_column_indexes(headers, "code_name", "code")
    type_idx = _find_column_indexes(headers, "type")
    size_idx = _find_column_indexes(headers, "size")
    od_idx = _find_column_indexes(headers, "od_in")
    lbs_outer_idx = _find_column_indexes(headers, "lbs_kft_outer")
    lbs_inner_idx = _find_column_indexes(headers, "lbs_kft_inner")
    r25_idx = _find_column_indexes(headers, "r25")
    r75_idx = _find_column_indexes(headers, "r75")

    for row in rows[1:]:
        if not row:
            continue
        code = _get_cell(row, code_idx)
        ctype = _get_cell(row, type_idx)
        size = _get_cell(row, size_idx)
        if not code or not ctype:
            continue

        name = f"{size} {ctype} ({code})".strip() if size and size != "0" else f"{code} ({ctype})".strip()
        od_in = _num(_get_cell(row, od_idx), 0.0)
        lbs_outer = _num(_get_cell(row, lbs_outer_idx), 0.0)
        lbs_inner = _num(_get_cell(row, lbs_inner_idx), 0.0)
        r25 = _num(_get_cell(row, r25_idx), 0.0)
        r75 = _num(_get_cell(row, r75_idx), r25 * 1.202 if r25 > 0.0 else 0.0)
        if r75 > 999.9:
            r75 = r25 * 1.202

        ctype_upper = ctype.upper()
        if ctype_upper == "CU":
            cp = lbs_outer / 1000.0 * 192.0
        elif ctype_upper == "ACCC":
            cp = lbs_outer / 1000.0 * 433.0 + lbs_inner / 1000.0 * 369.0
        else:
            cp = lbs_outer / 1000.0 * 433.0 + lbs_inner / 1000.0 * 216.0

        lookup[name.strip().lower()] = {
            "od_in": od_in * 25.4,
            "r25_ohm_per_m": r25 / 1609.344 if r25 > 0.0 else 0.0,
            "r75_ohm_per_m": r75 / 1609.344 if r75 > 0.0 else 0.0,
            "heat_cap_ws_per_m_c": cp * 3.28084,
        }

    return lookup


def load_conductors(path: str) -> tuple[list[Conductor], list[Conductor]]:
    # Load conductor/static rows from file (or fallback samples).
    if not os.path.exists(path):
        return sample_conductors(), sample_statics()
    if _is_xlsx(path):
        phase = load_conductors_from_condata(path)
        return phase or sample_conductors(), sample_statics()

    thermal_lookup = load_thermal_conductor_lookup()
    delim = _sniff_delimiter(path)
    phase: list[Conductor] = []
    statics: list[Conductor] = []

    with open(path, "r", newline="", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f, delimiter=delim)
        rows = list(reader)
        if not rows:
            return sample_conductors(), sample_statics()

        headers = rows[0]
        static_idxs = _find_column_indexes(headers, "is_static", "static", "wire_type")
        display_name_idxs = _find_column_indexes(headers, "display_name", "name", "conductor", "condname", "description")
        code_idxs = _find_column_indexes(headers, "code_name", "code")
        type_idxs = _find_column_indexes(headers, "type")
        size_idxs = _find_column_indexes(headers, "size")
        r_idxs = _find_column_indexes(headers, "r", "r_ohm_per_mi", "resistance", "r60", "r(ohms/mi)")
        xl_idxs = _find_column_indexes(headers, "xl", "x", "xl_ohm_per_mi", "xl_ohms_mi", "xl(ohms/mi)")
        xc_idxs = _find_column_indexes(
            headers,
            "xc",
            "c",
            "xc_ohms_mi",
            "xc_mohm_mi",
            "xc_mohms_mi",
            "xl_mohms_mi",
            "xc(ohms/mi)",
            "xc(ohms-mi)",
            "xc(mohm-mi)",
        )
        gmr_idxs = _find_column_indexes(headers, "gmr", "gmr_ft", "gmr(ft)")
        rad_idxs = _find_column_indexes(headers, "radius", "rad", "radius_ft", "radius(ft)")
        rate_a_idxs = _find_column_indexes(headers, "ratea", "rate_a", "ampa", "ratea(a)")
        rate_b_idxs = _find_column_indexes(headers, "rateb", "rate_b", "ampb", "rateb(a)")
        rate_c_idxs = _find_column_indexes(headers, "ratec", "rate_c", "ampc", "ratec(a)")
        od_idxs = _find_column_indexes(headers, "od_in", "diameter", "od")
        lbs_outer_idxs = _find_column_indexes(headers, "lbs_kft_outer")
        lbs_inner_idxs = _find_column_indexes(headers, "lbs_kft_inner")
        r25_idxs = _find_column_indexes(headers, "r25")
        r75_idxs = _find_column_indexes(headers, "r75")

        for raw in rows[1:]:
            if not raw:
                continue
            row = [str(value).strip() for value in raw]

            # Shared conductor sheets can contain duplicate NAME columns. For
            # the Dominion file layout, the last "Name" column is the real
            # conductor/static name and the earlier "NAME" column contains an
            # internal label like "2-Jan" that should not drive the UI.
            name = _get_last_cell(row, display_name_idxs)
            alt_name = _get_cell(row, display_name_idxs)
            if not name:
                code = _get_cell(row, code_idxs)
                ctype = _get_cell(row, type_idxs)
                size = _get_cell(row, size_idxs)
                if size and size != "0":
                    name = f"{size} {ctype} ({code})".strip()
                else:
                    name = f"{code} ({ctype})".strip()
            if not name:
                continue

            aliases: list[str] = []
            if alt_name and alt_name.lower() != name.lower():
                aliases.append(alt_name)

            is_static = str(_get_cell(row, static_idxs, "0")).lower() in {
                "1",
                "true",
                "static",
            }

            r = _num(_get_cell(row, r_idxs), 0.08)
            xl = _num(_get_cell(row, xl_idxs), 0.35)
            xc = _num(_get_cell(row, xc_idxs), 0.20)
            gmr = _num(_get_cell(row, gmr_idxs), 0.02)
            rad = _num(_get_cell(row, rad_idxs), 0.04)

            rate_a = _num(_get_cell(row, rate_a_idxs), 600.0)
            rate_b = _num(_get_cell(row, rate_b_idxs), 700.0)
            rate_c = _num(_get_cell(row, rate_c_idxs), 800.0)
            has_table_ratings = any(
                _get_cell(row, idxs) != ""
                for idxs in (rate_a_idxs, rate_b_idxs, rate_c_idxs)
            )

            od_in = _num(_get_cell(row, od_idxs), 1.0)
            lbs_outer = _num(_get_cell(row, lbs_outer_idxs), 0.0)
            lbs_inner = _num(_get_cell(row, lbs_inner_idxs), 0.0)
            r25 = _num(_get_cell(row, r25_idxs), 0.00005)
            r75 = _num(_get_cell(row, r75_idxs), r25 * 1.202)
            if r75 > 999.9:
                r75 = r25 * 1.202

            # Unit conversion copied from original intent.
            od_mm = od_in * 25.4
            r25_m = r25 / 1609.344
            r75_m = r75 / 1609.344
            ctype = _get_cell(row, type_idxs).upper()
            if ctype == "CU":
                cp = lbs_outer / 1000.0 * 192.0
            elif ctype == "ACCC":
                cp = lbs_outer / 1000.0 * 433.0 + lbs_inner / 1000.0 * 369.0
            else:
                cp = lbs_outer / 1000.0 * 433.0 + lbs_inner / 1000.0 * 216.0
            heat_cap = cp * 3.28084

            conductor = Conductor(
                name=name,
                aliases=aliases,
                has_table_ratings=has_table_ratings,
                is_static_default=is_static,
                gmr_ft=gmr,
                radius_ft=rad,
                r_ohm_per_mi=r,
                xl_ohm_per_mi=xl,
                xc_mohm_mi=xc,
                rate_a=rate_a,
                rate_b=rate_b,
                rate_c=rate_c,
                od_in=od_mm,
                r25_ohm_per_m=r25_m,
                r75_ohm_per_m=r75_m,
                heat_cap_ws_per_m_c=heat_cap,
            )

            thermal = thermal_lookup.get(name.strip().lower())
            if thermal is not None:
                conductor.od_in = thermal["od_in"]
                conductor.r25_ohm_per_m = thermal["r25_ohm_per_m"]
                conductor.r75_ohm_per_m = thermal["r75_ohm_per_m"]
                conductor.heat_cap_ws_per_m_c = thermal["heat_cap_ws_per_m_c"]

            if is_static:
                statics.append(conductor)
            else:
                phase.append(conductor)

    if not phase:
        phase = sample_conductors()
    if not statics:
        statics = sample_statics()

    return phase, statics


def load_structures(path: str) -> list[Structure]:
    if not os.path.exists(path):
        return sample_structures()

    if _is_xlsx(path):
        rows = _rows_from_xlsx(path)
    else:
        delim = _sniff_delimiter(path)
        with open(path, "r", newline="", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=delim)
            rows = list(reader)

    if len(rows) <= 1:
        return sample_structures()

    headers = ["" if value is None else str(value) for value in rows[0]]
    name_idxs = _find_column_indexes(headers, "name", "structure", "structure_name")
    ax_idxs = _find_column_indexes(headers, "ax", "a_x")
    ay_idxs = _find_column_indexes(headers, "ay", "a_y")
    bx_idxs = _find_column_indexes(headers, "bx", "b_x")
    by_idxs = _find_column_indexes(headers, "by", "b_y")
    cx_idxs = _find_column_indexes(headers, "cx", "c_x", "xs")
    cy_idxs = _find_column_indexes(headers, "cy", "c_y")
    g1x_idxs = _find_column_indexes(headers, "g1x", "g1_x")
    g1y_idxs = _find_column_indexes(headers, "g1y", "g1_y")
    g2x_idxs = _find_column_indexes(headers, "g2x", "g2_x")
    g2y_idxs = _find_column_indexes(headers, "g2y", "g2_y")

    header_has_coordinates = all(
        indexes
        for indexes in (name_idxs, ax_idxs, ay_idxs, bx_idxs, by_idxs, cx_idxs, cy_idxs)
    )

    structs: list[Structure] = []
    for row in rows[1:]:
        row = ["" if value is None else str(value).strip() for value in row]
        if header_has_coordinates:
            name = _get_cell(row, name_idxs)
            x1 = _num(_get_cell(row, ax_idxs), 0.0)
            y1 = _num(_get_cell(row, ay_idxs), 0.0)
            x2 = _num(_get_cell(row, bx_idxs), 0.0)
            y2 = _num(_get_cell(row, by_idxs), 0.0)
            x3 = _num(_get_cell(row, cx_idxs), 0.0)
            y3 = _num(_get_cell(row, cy_idxs), 0.0)
            g1x = _num(_get_cell(row, g1x_idxs), 0.0)
            g1y = _num(_get_cell(row, g1y_idxs), 0.0)
            g2x = _num(_get_cell(row, g2x_idxs), 0.0)
            g2y = _num(_get_cell(row, g2y_idxs), 0.0)
        else:
            vals = [c for c in row if c != ""]
            if len(vals) < 7:
                continue
            name = vals[0]
            x1, y1, x2, y2, x3, y3 = map(_num, vals[1:7])
            g1x = _num(vals[7], 0.0) if len(vals) > 7 else 0.0
            g1y = _num(vals[8], 0.0) if len(vals) > 8 else 0.0
            g2x = _num(vals[9], 0.0) if len(vals) > 9 else 0.0
            g2y = _num(vals[10], 0.0) if len(vals) > 10 else 0.0

        if not name:
            continue

        structs.append(
            Structure(
                name=name,
                a=[Point(x1, y1), Point(x2, y2), Point(x3, y3)],
                g=[Point(g1x, g1y), Point(g2x, g2y)],
            )
        )

    return structs or sample_structures()


def by_name(items: Iterable[Conductor | Structure], name: str):
    low = name.strip().lower()
    for item in items:
        if item.name.strip().lower() == low:
            return item
        if isinstance(item, Conductor):
            for alias in item.aliases:
                if alias.strip().lower() == low:
                    return item
    for item in items:
        if low in item.name.strip().lower():
            return item
        if isinstance(item, Conductor):
            for alias in item.aliases:
                if low in alias.strip().lower():
                    return item
    return None


def sample_conductors() -> list[Conductor]:
    return [
        Conductor(
            "1272 ACSR (BITTERN)",
            r_ohm_per_mi=0.0832,
            xl_ohm_per_mi=0.378,
            xc_mohm_mi=0.0855,
            rate_a=1190,
            rate_b=1276.2,
            rate_c=1579.5,
            has_table_ratings=True,
            gmr_ft=0.0455,
            radius_ft=0.05604,
            od_in=1.345 * 25.4,
            r25_ohm_per_m=0.0759 / 1609.344,
            r75_ohm_per_m=0.0903 / 1609.344,
            heat_cap_ws_per_m_c=(1197.7 / 1000.0 * 433.0 + 233.9 / 1000.0 * 216.0) * 3.28084,
        ),
        Conductor("1033.5 ACCC (ACCC)", 
            r_ohm_per_mi=0.028, 
            xl_ohm_per_mi=0.28, 
            xc_mohm_mi=0.20, 
            rate_a=1500, 
            rate_b=1650, 
            rate_c=1750, 
            has_table_ratings=True,
            gmr_ft=0.039, 
            radius_ft=0.051, 
            od_in=1.45, 
            r25_ohm_per_m=1.8e-5, 
            r75_ohm_per_m=2.2e-5, 
            heat_cap_ws_per_m_c=870
        ),
        Conductor("1113 ACCR (ACCR)", 
            r_ohm_per_mi=0.026, 
            xl_ohm_per_mi=0.27, 
            xc_mohm_mi=0.20, 
            rate_a=1650, 
            rate_b=1800, 
            rate_c=1900, 
            has_table_ratings=True,
            gmr_ft=0.04, 
            radius_ft=0.052, 
            od_in=1.5, 
            r25_ohm_per_m=1.6e-5, 
            r75_ohm_per_m=2.1e-5, 
            heat_cap_ws_per_m_c=880
        ),
        Conductor("1113 ACSS (ACSS)", 
            r_ohm_per_mi=0.024, 
            xl_ohm_per_mi=0.27, 
            xc_mohm_mi=0.20, 
            rate_a=1500, 
            rate_b=1700, 
            rate_c=1800, 
            has_table_ratings=True,
            gmr_ft=0.04, 
            radius_ft=0.052, 
            od_in=1.5, 
            r25_ohm_per_m=1.5e-5, 
            r75_ohm_per_m=2.0e-5, 
            heat_cap_ws_per_m_c=890
        ),
    ]


def sample_statics() -> list[Conductor]:
    return [
        Conductor(
            "1/4 GALV (None)",
            is_static_default=True,
            r_ohm_per_mi=7.83,
            xl_ohm_per_mi=2.07,
            xc_mohm_mi=0.1244,
            rate_a=130,
            rate_b=130,
            rate_c=130,
            has_table_ratings=True,
            gmr_ft=0.0104,
            radius_ft=0.0103
        ),
        Conductor("7#8 AW (none)", 
            is_static_default=True, 
            r_ohm_per_mi=3.06, 
            xl_ohm_per_mi=0.749, 
            xc_mohm_mi=0.1226, 
            rate_a=190, 
            rate_b=190, 
            rate_c=190, 
            has_table_ratings=True,
            gmr_ft=0.0021, 
            radius_ft=0.016
            ),
        Conductor(
            "3/8 GALV (None)",
            is_static_default=True,
            r_ohm_per_mi=5.51,
            xl_ohm_per_mi=0.93,
            xc_mohm_mi=0.1244,
            rate_a=130,
            rate_b=130,
            rate_c=130,
            has_table_ratings=True,
            gmr_ft=0.0156,
            radius_ft=0.0130
        ),
    ]


def sample_structures() -> list[Structure]:
    return [
        Structure("BPV",
            a=[Point(-18, 50), 
               Point(0, 58), 
               Point(18, 50)], 
            g=[Point(0, 72), 
               Point(0, 0)]
        ),
        Structure("HFRAME", 
            a=[Point(-22, 45), 
               Point(0, 45), 
               Point(22, 45)], 
            g=[Point(0, 62), 
               Point(0, 0)]
        ),
    ]
