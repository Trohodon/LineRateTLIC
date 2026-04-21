from dataclasses import dataclass
from typing import Optional


@dataclass
class Conductor:
    family: str
    code_word: str
    size_kcmil: Optional[float] = None
    stranding: Optional[str] = None

    al_area_in2: Optional[float] = None
    total_area_in2: Optional[float] = None
    al_layers: Optional[int] = None

    al_strand_dia_in: Optional[float] = None
    steel_strand_dia_in: Optional[float] = None
    steel_core_dia_in: Optional[float] = None
    od_in: Optional[float] = None

    al_weight_lb_per_kft: Optional[float] = None
    steel_weight_lb_per_kft: Optional[float] = None
    total_weight_lb_per_kft: Optional[float] = None

    al_percent: Optional[float] = None
    steel_percent: Optional[float] = None
    rbs_klb: Optional[float] = None

    dc_res_20c_ohm_per_mile: Optional[float] = None
    ac_res_25c_ohm_per_mile: Optional[float] = None
    ac_res_50c_ohm_per_mile: Optional[float] = None
    ac_res_75c_ohm_per_mile: Optional[float] = None
    ac_res_200c_ohm_per_mile: Optional[float] = None
    ac_res_250c_ohm_per_mile: Optional[float] = None
    stdol: Optional[float] = None

    gmr_ft: Optional[float] = None
    xa_60hz_ohm_per_mile: Optional[float] = None
    capacitive_reactance: Optional[float] = None
    ampacity_75c_amp: Optional[float] = None

    name: Optional[str] = None
    emissivity: Optional[float] = None
    absorptivity: Optional[float] = None
    max_temp_c: Optional[float] = None
