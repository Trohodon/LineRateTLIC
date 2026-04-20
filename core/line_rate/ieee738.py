from __future__ import annotations

import math
from typing import Optional

from core.line_rate.solar_ieee738 import solar_heat_gain
from core.line_rate.conductor import Conductor


INCH_TO_FT = 1.0 / 12.0
FT_PER_M = 3.280839895013123
OHM_PER_MILE_TO_OHM_PER_FT = 1.0 / 5280.0
MPS_TO_FPS = FT_PER_M


def inch_to_foot(value_in: float) -> float:
    return value_in * INCH_TO_FT


def ohm_per_mile_to_ohm_per_ft(value: float) -> float:
    return value * OHM_PER_MILE_TO_OHM_PER_FT


def _linear_interp(p1: tuple[float, float], p2: tuple[float, float], x: float) -> float:
    x1, y1 = p1
    x2, y2 = p2
    if x2 == x1:
        return y1
    return y1 + (y2 - y1) * (x - x1) / (x2 - x1)


def resolve_resistance_ohm_per_mile(
    conductor: Conductor,
    target_temp_c: float,
    r25_override: Optional[float] = None,
    r75_override: Optional[float] = None,
    r200_override: Optional[float] = None,
) -> dict:
    if target_temp_c <= 100.0:
        if r25_override is not None and r75_override is not None:
            r_tc = _linear_interp((25.0, r25_override), (75.0, r75_override), target_temp_c)
            return {
                "mode": "interpolated_from_overrides",
                "r25_ohm_per_mile": r25_override,
                "r75_ohm_per_mile": r75_override,
                "r200_ohm_per_mile": r200_override,
                "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
                "r_tc_ohm_per_mile": r_tc,
            }

        conductor_r25 = conductor.ac_res_25c_ohm_per_mile
        conductor_r75 = conductor.ac_res_75c_ohm_per_mile
        if conductor_r25 is not None and conductor_r75 is not None and abs(conductor_r75 - conductor_r25) > 1e-12:
            r_tc = _linear_interp((25.0, conductor_r25), (75.0, conductor_r75), target_temp_c)
            return {
                "mode": "interpolated_from_conductor_data",
                "r25_ohm_per_mile": conductor_r25,
                "r75_ohm_per_mile": conductor_r75,
                "r200_ohm_per_mile": conductor.ac_res_200c_ohm_per_mile,
                "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
                "r_tc_ohm_per_mile": r_tc,
            }
    else:
        low_r = r25_override if r25_override is not None else conductor.ac_res_25c_ohm_per_mile

        high_temp = 200.0
        high_r = r200_override if r200_override is not None else conductor.ac_res_200c_ohm_per_mile
        if target_temp_c > 200.0 and conductor.ac_res_250c_ohm_per_mile is not None:
            high_temp = 250.0
            high_r = conductor.ac_res_250c_ohm_per_mile

        if low_r is not None and high_r is not None and abs(high_r - low_r) > 1e-12:
            r_tc = _linear_interp((25.0, low_r), (high_temp, high_r), target_temp_c)
            return {
                "mode": f"interpolated_using_25_{int(high_temp)}",
                "r25_ohm_per_mile": low_r,
                "r75_ohm_per_mile": r75_override if r75_override is not None else conductor.ac_res_75c_ohm_per_mile,
                "r200_ohm_per_mile": r200_override if r200_override is not None else conductor.ac_res_200c_ohm_per_mile,
                "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
                "r_tc_ohm_per_mile": r_tc,
            }

        if r25_override is not None and r75_override is not None:
            r_tc = _linear_interp((25.0, r25_override), (75.0, r75_override), target_temp_c)
            return {
                "mode": "extrapolated_from_overrides_gt100",
                "r25_ohm_per_mile": r25_override,
                "r75_ohm_per_mile": r75_override,
                "r200_ohm_per_mile": r200_override,
                "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
                "r_tc_ohm_per_mile": r_tc,
            }

        conductor_r25 = conductor.ac_res_25c_ohm_per_mile
        conductor_r75 = conductor.ac_res_75c_ohm_per_mile
        if conductor_r25 is not None and conductor_r75 is not None and abs(conductor_r75 - conductor_r25) > 1e-12:
            r_tc = _linear_interp((25.0, conductor_r25), (75.0, conductor_r75), target_temp_c)
            return {
                "mode": "extrapolated_from_conductor_data_gt100",
                "r25_ohm_per_mile": conductor_r25,
                "r75_ohm_per_mile": conductor_r75,
                "r200_ohm_per_mile": conductor.ac_res_200c_ohm_per_mile,
                "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
                "r_tc_ohm_per_mile": r_tc,
            }

    single_r = conductor.ac_res_25c_ohm_per_mile
    if single_r is None:
        single_r = conductor.dc_res_20c_ohm_per_mile

    if single_r is None:
        raise ValueError(
            f"Conductor '{conductor.code_word}' has no usable resistance data. "
            f"Provide R25/R75 or add temperature-specific resistance values to the workbook."
        )

    return {
        "mode": "single_r_direct",
        "r25_ohm_per_mile": conductor.ac_res_25c_ohm_per_mile,
        "r75_ohm_per_mile": conductor.ac_res_75c_ohm_per_mile,
        "r200_ohm_per_mile": conductor.ac_res_200c_ohm_per_mile,
        "r250_ohm_per_mile": conductor.ac_res_250c_ohm_per_mile,
        "r_tc_ohm_per_mile": single_r,
    }


def mean_film_temperature(ts_c: float, ta_c: float) -> float:
    return (ts_c + ta_c) / 2.0


def air_dynamic_viscosity_lb_per_ft_s(tfilm_c: float) -> float:
    return (9.806e-7 * (tfilm_c + 273.0) ** 1.5) / (tfilm_c + 383.4)


def air_density_lb_per_ft3(tfilm_c: float, elevation_ft: float) -> float:
    numerator = 0.080695 - 2.901e-6 * elevation_ft + 3.7e-11 * (elevation_ft ** 2)
    denominator = 1.0 + 0.00367 * tfilm_c
    return numerator / denominator


def air_thermal_conductivity_w_per_ft_c(tfilm_c: float) -> float:
    return 7.388e-3 + 2.279e-5 * tfilm_c - 1.343e-9 * (tfilm_c ** 2)


def wind_direction_factor_from_beta(beta_deg: float) -> float:
    beta = math.radians(beta_deg)
    return 1.194 - math.sin(beta) - 0.194 * math.cos(2.0 * beta) + 0.368 * math.sin(2.0 * beta)


def reynolds_number(
    diameter_ft: float,
    rho_f_lb_per_ft3: float,
    wind_fps: float,
    mu_f_lb_per_ft_s: float,
) -> float:
    return diameter_ft * rho_f_lb_per_ft3 * wind_fps / mu_f_lb_per_ft_s


def natural_convection_loss_w_per_ft(ts_c: float, ta_c: float, diameter_ft: float, rho_f_lb_per_ft3: float) -> float:
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0:
        return 0.0
    return 1.825 * (rho_f_lb_per_ft3 ** 0.5) * (diameter_ft ** 0.75) * (delta_t ** 1.25)


def forced_convection_losses_w_per_ft(
    ts_c: float,
    ta_c: float,
    diameter_ft: float,
    wind_fps: float,
    beta_deg: float,
    rho_f_lb_per_ft3: float,
    mu_f_lb_per_ft_s: float,
    k_f_w_per_ft_c: float,
) -> tuple[float, float]:
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0 or wind_fps <= 0.0:
        return 0.0, 0.0

    k_angle = wind_direction_factor_from_beta(beta_deg)
    n_re = reynolds_number(diameter_ft, rho_f_lb_per_ft3, wind_fps, mu_f_lb_per_ft_s)

    qc1 = k_angle * (1.01 + 1.35 * (n_re ** 0.52)) * k_f_w_per_ft_c * delta_t
    qc2 = k_angle * 0.754 * (n_re ** 0.60) * k_f_w_per_ft_c * delta_t
    return qc1, qc2


def convection_loss(
    ts_c: float,
    ta_c: float,
    diameter_ft: float,
    wind_speed_mps: float,
    wind_angle_to_axis_deg: float,
    elevation_m: float,
) -> dict:
    tfilm = mean_film_temperature(ts_c, ta_c)
    wind_fps = wind_speed_mps * MPS_TO_FPS
    elevation_ft = elevation_m * FT_PER_M

    mu_f = air_dynamic_viscosity_lb_per_ft_s(tfilm)
    rho_f = air_density_lb_per_ft3(tfilm, elevation_ft)
    k_f = air_thermal_conductivity_w_per_ft_c(tfilm)

    # Boss workbook uses beta = angle to the perpendicular.
    beta_deg = 90.0 - wind_angle_to_axis_deg

    qcn = natural_convection_loss_w_per_ft(ts_c, ta_c, diameter_ft, rho_f)
    qc1, qc2 = forced_convection_losses_w_per_ft(ts_c, ta_c, diameter_ft, wind_fps, beta_deg, rho_f, mu_f, k_f)
    qc = max(qcn, qc1, qc2)

    n_re = reynolds_number(diameter_ft, rho_f, wind_fps, mu_f) if wind_fps > 0 else 0.0

    return {
        "qc_w_per_ft": qc,
        "qcn_w_per_ft": qcn,
        "qc1_w_per_ft": qc1,
        "qc2_w_per_ft": qc2,
        "tfilm_c": tfilm,
        "rho_f_lb_per_ft3": rho_f,
        "mu_f_lb_per_ft_s": mu_f,
        "k_f_w_per_ft_c": k_f,
        "n_re": n_re,
        "k_angle": wind_direction_factor_from_beta(beta_deg),
        "beta_deg": beta_deg,
        "wind_fps": wind_fps,
        "elevation_ft": elevation_ft,
        "delta_t_c": max(ts_c - ta_c, 0.0),
    }


def radiated_heat_loss(ts_c: float, ta_c: float, diameter_ft: float, emissivity: float) -> dict:
    qr_w_per_ft = 1.656 * diameter_ft * emissivity * ((((ts_c + 273.0) / 100.0) ** 4) - (((ta_c + 273.0) / 100.0) ** 4))
    return {
        "qr_w_per_ft": qr_w_per_ft,
    }


def calculate_steady_state_rating(
    conductor: Conductor,
    ambient_temp_c: float,
    wind_speed_mps: float,
    wind_angle_deg: float,
    elevation_m: float,
    target_temp_c: float,
    emissivity: Optional[float] = None,
    absorptivity: Optional[float] = None,
    latitude_deg: Optional[float] = None,
    line_azimuth_deg: Optional[float] = None,
    input_date=None,
    input_time=None,
    atmosphere_type: str = "clear",
    r25_override: Optional[float] = None,
    r75_override: Optional[float] = None,
    r200_override: Optional[float] = None,
) -> dict:
    if conductor.od_in is None:
        raise ValueError(f"Conductor '{conductor.code_word}' is missing OD_IN.")

    if latitude_deg is None or line_azimuth_deg is None or input_date is None or input_time is None:
        raise ValueError("Latitude, line azimuth, date, and time are required for the full IEEE 738 solar model.")

    eps = emissivity if emissivity is not None else (conductor.emissivity if conductor.emissivity is not None else 0.5)
    alpha = absorptivity if absorptivity is not None else (conductor.absorptivity if conductor.absorptivity is not None else 0.5)

    diameter_in = conductor.od_in
    diameter_ft = inch_to_foot(diameter_in)

    resistance_info = resolve_resistance_ohm_per_mile(
        conductor=conductor,
        target_temp_c=target_temp_c,
        r25_override=r25_override,
        r75_override=r75_override,
        r200_override=r200_override,
    )
    resistance_ohm_per_mile = resistance_info["r_tc_ohm_per_mile"]
    resistance_ohm_per_ft = ohm_per_mile_to_ohm_per_ft(resistance_ohm_per_mile)

    convection = convection_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_ft=diameter_ft,
        wind_speed_mps=wind_speed_mps,
        wind_angle_to_axis_deg=wind_angle_deg,
        elevation_m=elevation_m,
    )

    radiation = radiated_heat_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_ft=diameter_ft,
        emissivity=eps,
    )

    solar = solar_heat_gain(
        absorptivity=alpha,
        diameter_ft=diameter_ft,
        latitude_deg=latitude_deg,
        line_azimuth_deg=line_azimuth_deg,
        input_date=input_date,
        input_time=input_time,
        elevation_m=elevation_m,
        atmosphere_type=atmosphere_type,
    )

    net_ft = convection["qc_w_per_ft"] + radiation["qr_w_per_ft"] - solar["qs_w_per_ft"]
    amps = math.sqrt(net_ft / resistance_ohm_per_ft) if net_ft > 0.0 and resistance_ohm_per_ft > 0.0 else 0.0

    return {
        "code_word": conductor.code_word,
        "target_temp_c": target_temp_c,
        "ambient_temp_c": ambient_temp_c,
        "wind_speed_mps": wind_speed_mps,
        "wind_speed_fps": convection["wind_fps"],
        "wind_angle_deg": wind_angle_deg,
        "beta_deg": convection["beta_deg"],
        "elevation_m": elevation_m,
        "elevation_ft": convection["elevation_ft"],
        "diameter_ft": diameter_ft,
        "diameter_in": conductor.od_in,
        "resistance_mode": resistance_info["mode"],
        "r25_used_ohm_per_mile": resistance_info["r25_ohm_per_mile"],
        "r75_used_ohm_per_mile": resistance_info["r75_ohm_per_mile"],
        "r200_used_ohm_per_mile": resistance_info["r200_ohm_per_mile"],
        "r250_used_ohm_per_mile": resistance_info["r250_ohm_per_mile"],
        "resistance_ohm_per_mile": resistance_ohm_per_mile,
        "resistance_ohm_per_ft": resistance_ohm_per_ft,
        "qc_w_per_ft": convection["qc_w_per_ft"],
        "qcn_w_per_ft": convection["qcn_w_per_ft"],
        "qc1_w_per_ft": convection["qc1_w_per_ft"],
        "qc2_w_per_ft": convection["qc2_w_per_ft"],
        "qr_w_per_ft": radiation["qr_w_per_ft"],
        "qs_w_per_ft": solar["qs_w_per_ft"],
        "amps": amps,
        "tfilm_c": convection["tfilm_c"],
        "rho_f_lb_per_ft3": convection["rho_f_lb_per_ft3"],
        "mu_f_lb_per_ft_s": convection["mu_f_lb_per_ft_s"],
        "k_f_w_per_ft_c": convection["k_f_w_per_ft_c"],
        "n_re": convection["n_re"],
        "k_angle": convection["k_angle"],
        "emissivity": eps,
        "absorptivity": alpha,
        "solar": solar,
    }
