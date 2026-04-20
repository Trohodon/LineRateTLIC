from __future__ import annotations

import math
from datetime import datetime, date, time


FT_PER_M = 3.280839895013123
W_PER_FT2_TO_W_PER_M2 = 10.763910416709722


CLEAR_AIR_COEFFICIENTS_US = {
    "A": -3.9241,
    "B": 5.9276,
    "C": -0.17856,
    "D": 0.003223,
    "E": -3.3549e-5,
    "F": 1.8053e-7,
    "G": -3.7868e-10,
}

INDUSTRIAL_AIR_COEFFICIENTS_US = {
    "A": 4.9408,
    "B": 1.3202,
    "C": 0.061444,
    "D": -0.0029411,
    "E": 5.07752e-5,
    "F": -4.03627e-7,
    "G": 1.22967e-9,
}


def _sin_deg(value: float) -> float:
    return math.sin(math.radians(value))


def _cos_deg(value: float) -> float:
    return math.cos(math.radians(value))


def _tan_deg(value: float) -> float:
    return math.tan(math.radians(value))


def _asin_deg(value: float) -> float:
    value = max(-1.0, min(1.0, value))
    return math.degrees(math.asin(value))


def _atan_deg(value: float) -> float:
    return math.degrees(math.atan(value))


def parse_date_input(date_text: str) -> date:
    raw = date_text.strip()

    fmts = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%d",
        "%m-%d-%Y",
        "%m-%d-%y",
    ]

    for fmt in fmts:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass

    raise ValueError(
        f"Invalid date '{date_text}'. Use MM/DD/YYYY, MM/DD/YY, or YYYY-MM-DD."
    )


def parse_time_input(time_text: str) -> time:
    raw = time_text.strip().upper()

    fmts = [
        "%H:%M",
        "%H:%M:%S",
        "%I:%M %p",
        "%I:%M:%S %p",
        "%I %p",
    ]

    for fmt in fmts:
        try:
            return datetime.strptime(raw, fmt).time()
        except ValueError:
            pass

    raise ValueError(
        f"Invalid time '{time_text}'. Use HH:MM, HH:MM:SS, or 11:00 AM style."
    )


def day_of_year(input_date: date) -> int:
    return input_date.timetuple().tm_yday


def decimal_hour(input_time: time) -> float:
    return input_time.hour + (input_time.minute / 60.0) + (input_time.second / 3600.0)


def hour_angle_deg(local_time_hours: float) -> float:
    return (local_time_hours - 12.0) * 15.0


def solar_declination_deg(day_num: int) -> float:
    return 23.45 * _sin_deg(((284.0 + day_num) / 365.0) * 360.0)


def solar_altitude_deg(latitude_deg: float, declination_deg: float, hour_angle_deg_value: float) -> float:
    hc = _asin_deg(
        _cos_deg(latitude_deg) * _cos_deg(declination_deg) * _cos_deg(hour_angle_deg_value)
        + _sin_deg(latitude_deg) * _sin_deg(declination_deg)
    )
    return max(hc, 0.0)


def solar_azimuth_variable(latitude_deg: float, declination_deg: float, hour_angle_deg_value: float) -> float:
    denominator = (
        _sin_deg(latitude_deg) * _cos_deg(hour_angle_deg_value)
        - _cos_deg(latitude_deg) * _tan_deg(declination_deg)
    )

    if abs(denominator) < 1e-12:
        return math.copysign(float("inf"), _sin_deg(hour_angle_deg_value))

    return _sin_deg(hour_angle_deg_value) / denominator


def solar_azimuth_constant(hour_angle_deg_value: float, chi: float) -> float:
    if -180.0 <= hour_angle_deg_value < 0.0:
        return 0.0 if chi >= 0.0 else 180.0
    if 0.0 <= hour_angle_deg_value <= 180.0:
        return 180.0 if chi >= 0.0 else 360.0
    return 180.0 if chi >= 0.0 else 360.0


def solar_azimuth_deg(latitude_deg: float, declination_deg: float, hour_angle_deg_value: float) -> tuple[float, float, float]:
    chi = solar_azimuth_variable(latitude_deg, declination_deg, hour_angle_deg_value)
    c = solar_azimuth_constant(hour_angle_deg_value, chi)
    zc = c + _atan_deg(chi)
    return zc, chi, c


def angle_of_incidence_deg(solar_altitude_deg_value: float, solar_azimuth_deg_value: float, line_azimuth_deg: float) -> float:
    cosine_term = _cos_deg(solar_altitude_deg_value) * _cos_deg(solar_azimuth_deg_value - line_azimuth_deg)
    cosine_term = max(-1.0, min(1.0, cosine_term))
    return math.degrees(math.acos(cosine_term))


def sea_level_solar_intensity_w_per_ft2(solar_altitude_deg_value: float, atmosphere_type: str) -> float:
    atmosphere = atmosphere_type.strip().lower()
    coeffs = CLEAR_AIR_COEFFICIENTS_US if atmosphere == "clear" else INDUSTRIAL_AIR_COEFFICIENTS_US

    hc = solar_altitude_deg_value
    qs = (
        coeffs["A"]
        + coeffs["B"] * hc
        + coeffs["C"] * (hc ** 2)
        + coeffs["D"] * (hc ** 3)
        + coeffs["E"] * (hc ** 4)
        + coeffs["F"] * (hc ** 5)
        + coeffs["G"] * (hc ** 6)
    )
    return max(qs, 0.0)


def solar_elevation_correction_from_m(elevation_m: float) -> float:
    return 1.0 + 1.148e-4 * elevation_m - 1.108e-8 * (elevation_m ** 2)


def solar_heat_gain(
    absorptivity: float,
    diameter_ft: float,
    latitude_deg: float,
    line_azimuth_deg: float,
    input_date: date,
    input_time: time,
    elevation_m: float,
    atmosphere_type: str,
) -> dict:
    n = day_of_year(input_date)
    time_hours = decimal_hour(input_time)
    omega = hour_angle_deg(time_hours)
    delta = solar_declination_deg(n)
    hc = solar_altitude_deg(latitude_deg, delta, omega)

    ksolar = solar_elevation_correction_from_m(elevation_m)

    if hc <= 0.0:
        return {
            "qs_w_per_ft": 0.0,
            "qs_w_per_m": 0.0,
            "n_day": n,
            "hour_decimal": time_hours,
            "omega_deg": omega,
            "delta_deg": delta,
            "hc_deg": 0.0,
            "zc_deg": 0.0,
            "chi": 0.0,
            "c_constant": 0.0,
            "theta_deg": 0.0,
            "qs_sea_level_w_per_ft2": 0.0,
            "qs_sea_level_w_per_m2": 0.0,
            "qse_w_per_ft2": 0.0,
            "qse_w_per_m2": 0.0,
            "ksolar": ksolar,
            "projected_area_ft2_per_ft": diameter_ft,
            "projected_area_m2_per_m": diameter_ft / FT_PER_M,
            "atmosphere_type": atmosphere_type,
        }

    zc, chi, c_constant = solar_azimuth_deg(latitude_deg, delta, omega)
    theta = angle_of_incidence_deg(hc, zc, line_azimuth_deg)

    qs_sea_level_ft2 = sea_level_solar_intensity_w_per_ft2(hc, atmosphere_type)
    qse_ft2 = qs_sea_level_ft2 * ksolar
    qs_w_per_ft = absorptivity * qse_ft2 * _sin_deg(theta) * diameter_ft

    return {
        "qs_w_per_ft": max(qs_w_per_ft, 0.0),
        "qs_w_per_m": max(qs_w_per_ft, 0.0) * FT_PER_M,
        "n_day": n,
        "hour_decimal": time_hours,
        "omega_deg": omega,
        "delta_deg": delta,
        "hc_deg": hc,
        "zc_deg": zc,
        "chi": chi,
        "c_constant": c_constant,
        "theta_deg": theta,
        "qs_sea_level_w_per_ft2": qs_sea_level_ft2,
        "qs_sea_level_w_per_m2": qs_sea_level_ft2 * W_PER_FT2_TO_W_PER_M2,
        "qse_w_per_ft2": qse_ft2,
        "qse_w_per_m2": qse_ft2 * W_PER_FT2_TO_W_PER_M2,
        "ksolar": ksolar,
        "projected_area_ft2_per_ft": diameter_ft,
        "projected_area_m2_per_m": diameter_ft / FT_PER_M,
        "atmosphere_type": atmosphere_type,
    }
