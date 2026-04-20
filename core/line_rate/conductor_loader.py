from __future__ import annotations

from typing import Dict, List, Optional

from core.line_rate.xlsx_reader import list_sheet_names, read_sheet_table
from core.line_rate.conductor import Conductor


class ConductorDatabase:
    def __init__(self) -> None:
        self.by_family: Dict[str, List[Conductor]] = {}
        self.source_path: Optional[str] = None

    def add_family(self, family: str, conductors: List[Conductor]) -> None:
        self.by_family[family] = conductors

    def get_families(self) -> List[str]:
        return sorted(self.by_family.keys())

    def get_conductors(self, family: str) -> List[Conductor]:
        return self.by_family.get(family, [])

    def find_conductor(self, family: str, code_word: str) -> Optional[Conductor]:
        for conductor in self.get_conductors(family):
            if str(conductor.code_word).strip().upper() == str(code_word).strip().upper():
                return conductor
        return None


def _normalize_column_name(col_name: str) -> str:
    name = str(col_name).strip().upper()
    replacements = {
        "\n": "_",
        " ": "_",
        "-": "_",
        "/": "_",
        "(": "",
        ")": "",
        ".": "",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name


def _is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _to_float(value) -> Optional[float]:
    if _is_blank(value):
        return None
    if isinstance(value, str) and value.strip().upper() in {"N/A", "NA", "N\\A", "Ν/Α"}:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_str(value) -> Optional[str]:
    if _is_blank(value):
        return None
    if isinstance(value, str) and value.strip().upper() in {"N/A", "NA", "N\\A", "Ν/Α"}:
        return None
    return str(value).strip()


def _to_float_unless_sentinel(value, sentinel: float) -> Optional[float]:
    number = _to_float(value)
    if number is None:
        return None
    if abs(number - sentinel) < 1e-12:
        return None
    return number


def _table_to_row_maps(table: list[list[str | None]]) -> list[dict[str, list[str | None]]]:
    if not table:
        return []

    header_row = table[0]
    normalized_headers = [_normalize_column_name(value or "") for value in header_row]

    rows: list[dict[str, list[str | None]]] = []
    for raw_row in table[1:]:
        row_map: dict[str, list[str | None]] = {}
        has_value = False

        for index, header in enumerate(normalized_headers):
            if header == "":
                continue

            value = raw_row[index] if index < len(raw_row) else None
            if not _is_blank(value):
                has_value = True
            row_map.setdefault(header, []).append(value)

        if has_value:
            rows.append(row_map)

    return rows


def _get_all_present(row: dict[str, list[str | None]], name: str) -> list[str | None]:
    values = row.get(name, [])
    return [value for value in values if not _is_blank(value)]


def _get_first_present(row: dict[str, list[str | None]], possible_names: list[str]):
    for name in possible_names:
        values = _get_all_present(row, name)
        if values:
            return values[0]
    return None


def _looks_like_conductordata_workbook(header_row: list[str | None]) -> bool:
    cols = {_normalize_column_name(value or "") for value in header_row if not _is_blank(value)}
    required = {"TYPE", "CODE", "RADIUSFT", "GMRFT"}
    resistance_aliases = {"ROHMS_MI", "ROHMS_M", "R"}
    rate_aliases = {"RATEAA", "RATEBA", "RATECA"}
    return required.issubset(cols) and bool(cols & resistance_aliases) and rate_aliases.issubset(cols)


def _looks_like_condata_workbook(header_row: list[str | None]) -> bool:
    cols = {_normalize_column_name(value or "") for value in header_row if not _is_blank(value)}
    required = {"CODE_NAME", "TYPE", "OD_IN", "R25", "R75", "NAME"}
    return required.issubset(cols)


def _build_conductor_from_conductordata_row(sheet_name: str, row: dict[str, list[str | None]]) -> Optional[Conductor]:
    code_word = _to_str(_get_first_present(row, ["CODE"]))
    if code_word is None:
        return None

    family = _to_str(_get_first_present(row, ["TYPE"])) or sheet_name

    all_name_values = _get_all_present(row, "NAME")
    size_from_first_name = _to_float(all_name_values[0]) if all_name_values else None
    pretty_name = _to_str(all_name_values[1]) if len(all_name_values) >= 2 else None

    radius_ft = _to_float(_get_first_present(row, ["RADIUSFT"]))
    od_in = radius_ft * 24.0 if radius_ft is not None else None

    resistance = _to_float(_get_first_present(row, ["ROHMS_MI", "ROHMS_M", "R"]))
    gmr_ft = _to_float(_get_first_present(row, ["GMRFT"]))

    rate_a = _to_float(_get_first_present(row, ["RATEAA"]))
    rate_b = _to_float(_get_first_present(row, ["RATEBA"]))
    rate_c = _to_float(_get_first_present(row, ["RATECA"]))

    return Conductor(
        family=family,
        code_word=code_word,
        size_kcmil=size_from_first_name,
        stranding=None,
        al_area_in2=None,
        total_area_in2=None,
        al_layers=None,
        al_strand_dia_in=None,
        steel_strand_dia_in=None,
        steel_core_dia_in=None,
        od_in=od_in,
        al_weight_lb_per_kft=None,
        steel_weight_lb_per_kft=None,
        total_weight_lb_per_kft=None,
        al_percent=None,
        steel_percent=None,
        rbs_klb=None,
        dc_res_20c_ohm_per_mile=resistance,
        ac_res_25c_ohm_per_mile=resistance,
        ac_res_50c_ohm_per_mile=resistance,
        ac_res_75c_ohm_per_mile=resistance,
        ac_res_200c_ohm_per_mile=None,
        ac_res_250c_ohm_per_mile=None,
        stdol=None,
        gmr_ft=gmr_ft,
        xa_60hz_ohm_per_mile=_to_float(_get_first_present(row, ["XLOHMS_MI", "XLOHMS_R", "XLOHMS_R_", "XL"])),
        capacitive_reactance=_to_float(_get_first_present(row, ["XCOHMS_MI", "XCOHMS_T", "XC"])),
        ampacity_75c_amp=rate_b if rate_b is not None else (rate_a if rate_a is not None else rate_c),
        name=pretty_name or code_word,
        emissivity=None,
        absorptivity=None,
        max_temp_c=None,
    )


def _build_conductor_from_condata_row(sheet_name: str, row: dict[str, list[str | None]]) -> Optional[Conductor]:
    code_word = _to_str(_get_first_present(row, ["CODE_NAME"]))
    if code_word is None:
        return None

    family = _to_str(_get_first_present(row, ["TYPE"])) or sheet_name
    od_in = _to_float(_get_first_present(row, ["OD_IN"]))
    area_sq_in = _to_float(_get_first_present(row, ["AREA_SQIN"]))
    dc_r20 = _to_float(_get_first_present(row, ["DC_R20"]))
    r25 = _to_float(_get_first_present(row, ["R25"]))
    r50 = _to_float(_get_first_present(row, ["R50"]))
    r75 = _to_float(_get_first_present(row, ["R75"]))
    r200 = _to_float(_get_first_present(row, ["R200"]))
    r250 = _to_float(_get_first_present(row, ["R250"]))
    stdol = _to_float(_get_first_present(row, ["STDOL"]))
    gmr_ft = _to_float(_get_first_present(row, ["GMR_FT", "GMRFT"]))
    xl_ohm_per_mile = _to_float(_get_first_present(row, ["XL_OHMS_MI", "XLOHMS_MI", "XL"]))
    xc_mohm_mile = _to_float(_get_first_present(row, ["XC_MOHMS_MI", "XL_MOHMS_MI", "XCMOHMS_MI", "XC"]))

    return Conductor(
        family=family,
        code_word=code_word,
        size_kcmil=_to_float(_get_first_present(row, ["SIZE"])),
        stranding=_to_str(_get_first_present(row, ["STRAND"])),
        al_area_in2=area_sq_in,
        total_area_in2=area_sq_in,
        al_layers=None,
        al_strand_dia_in=_to_float(_get_first_present(row, ["DIAM_OUTERIN"])),
        steel_strand_dia_in=_to_float_unless_sentinel(_get_first_present(row, ["DIAM_INNERIN"]), 9999.0),
        steel_core_dia_in=None,
        od_in=od_in,
        al_weight_lb_per_kft=_to_float(_get_first_present(row, ["LBS_KFT_OUTER"])),
        steel_weight_lb_per_kft=_to_float(_get_first_present(row, ["LBS_KFT_INNER"])),
        total_weight_lb_per_kft=None,
        al_percent=None,
        steel_percent=None,
        rbs_klb=None if _to_float(_get_first_present(row, ["UTS_LBS"])) is None else _to_float(_get_first_present(row, ["UTS_LBS"])) / 1000.0,
        dc_res_20c_ohm_per_mile=dc_r20,
        ac_res_25c_ohm_per_mile=r25,
        ac_res_50c_ohm_per_mile=r50,
        ac_res_75c_ohm_per_mile=r75,
        ac_res_200c_ohm_per_mile=r200,
        ac_res_250c_ohm_per_mile=r250,
        stdol=stdol,
        gmr_ft=gmr_ft,
        xa_60hz_ohm_per_mile=xl_ohm_per_mile,
        capacitive_reactance=xc_mohm_mile,
        ampacity_75c_amp=None,
        name=_to_str(_get_first_present(row, ["NAME"])) or code_word,
        emissivity=None,
        absorptivity=None,
        max_temp_c=None,
    )


def load_conductor_database(filepath: str) -> ConductorDatabase:
    database = ConductorDatabase()
    database.source_path = filepath

    for sheet_name in list_sheet_names(filepath):
        table = read_sheet_table(filepath, sheet_name)
        if not table:
            continue

        grouped: Dict[str, List[Conductor]] = {}
        row_maps = _table_to_row_maps(table)

        if _looks_like_conductordata_workbook(table[0]):
            builder = _build_conductor_from_conductordata_row
        elif _looks_like_condata_workbook(table[0]):
            builder = _build_conductor_from_condata_row
        else:
            continue

        for row in row_maps:
            conductor = builder(sheet_name, row)
            if conductor is None:
                continue
            grouped.setdefault(conductor.family, []).append(conductor)

        for family_name, family_conductors in grouped.items():
            existing = database.get_conductors(family_name)
            database.add_family(family_name, existing + family_conductors)

    return database
