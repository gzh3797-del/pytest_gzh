"""
功率计算模块（独立，逻辑来自 power_calculate.py，不依赖原工程）
"""
import math
import cmath


def active_power(voltage: float, current: float, v_angle_deg: float, i_angle_deg: float) -> float:
    """有功功率 P = U × I × cos(θ_V − θ_I)，单位 kW"""
    angle = v_angle_deg - i_angle_deg
    return voltage * current * math.cos(math.radians(angle)) / 1000


def reactive_power(voltage: float, current: float, v_angle_deg: float, i_angle_deg: float) -> float:
    """无功功率 Q = U × I × sin(θ_V − θ_I)，单位 kvar"""
    angle = v_angle_deg - i_angle_deg
    return voltage * current * math.sin(math.radians(angle)) / 1000


def apparent_power(voltage: float, current: float) -> float:
    """视在功率 S = U × I，单位 kVA"""
    return voltage * current / 1000


def power_factor(v_angle_deg: float, i_angle_deg: float) -> float:
    """功率因数 PF = cos(θ_V − θ_I)"""
    return math.cos(math.radians(v_angle_deg - i_angle_deg))


def phase_power(v: float, i: float, v_ang: float, i_ang: float) -> tuple[float, float, float]:
    """单相 (P, Q, S)"""
    return (
        active_power(v, i, v_ang, i_ang),
        reactive_power(v, i, v_ang, i_ang),
        apparent_power(v, i),
    )


def line_to_line_voltage(ua, ub, uc, ua_p, ub_p, uc_p) -> tuple[float, float, float]:
    """由相电压和相角计算线电压 Uab, Ubc, Uca"""
    va = cmath.rect(ua, math.radians(ua_p))
    vb = cmath.rect(ub, math.radians(ub_p))
    vc = cmath.rect(uc, math.radians(uc_p))
    uab = abs(va - vb)
    ubc = abs(vb - vc)
    uca = abs(vc - va)
    return round(uab, 5), round(ubc, 5), round(uca, 5)


def sys_power(*phase_powers: float) -> float:
    """系统功率 = 各相之和"""
    return sum(phase_powers)


def calc_2e3wd_power(ua, ub, uc, ua_p, ub_p, uc_p, ia, ib, ic, ia_p, ib_p, ic_p
                     ) -> tuple[float, float, float]:
    """
    2E3WD Aron 两表法系统功率
    P_sys = Uab·Ia·cos(θ_Uab−θ_Ia) + Ubc·Ic·cos(θ_Ubc−θ_Ic)
    """
    va = cmath.rect(ua, math.radians(ua_p))
    vb = cmath.rect(ub, math.radians(ub_p))
    vc = cmath.rect(uc, math.radians(uc_p))
    uab_vec = va - vb
    ubc_vec = vb - vc
    uab = abs(uab_vec)
    ubc = abs(ubc_vec)
    uab_p = math.degrees(cmath.phase(uab_vec))
    ubc_p = math.degrees(cmath.phase(ubc_vec))

    p_sys = (uab * ia * math.cos(math.radians(uab_p - ia_p)) +
             ubc * ic * math.cos(math.radians(ubc_p - ic_p))) / 1000
    q_sys = (uab * ia * math.sin(math.radians(uab_p - ia_p)) +
             ubc * ic * math.sin(math.radians(ubc_p - ic_p))) / 1000
    s_sys = math.sqrt(p_sys ** 2 + q_sys ** 2)
    return round(p_sys, 5), round(q_sys, 5), round(s_sys, 5)


def calc_3e4wd_power(ua, ub, uc, ua_p, ub_p, uc_p, ia, ib, ic, ia_p, ib_p, ic_p
                     ) -> tuple[float, float, float]:
    """
    3E3WD 系统功率 — Aron 两表法（与参考脚本 calculate_3e4wd_power 一致）
    S = Vab·Ia* + Vcb·Ic*，单位 kW/kvar/kVA
    """
    va = cmath.rect(ua, math.radians(ua_p))
    vb = cmath.rect(ub, math.radians(ub_p))
    vc = cmath.rect(uc, math.radians(uc_p))
    ia_c = cmath.rect(ia, math.radians(ia_p))
    ic_c = cmath.rect(ic, math.radians(ic_p))
    vab = va - vb
    vcb = vc - vb
    s_total = vab * ia_c.conjugate() + vcb * ic_c.conjugate()
    p_sys = s_total.real / 1000
    q_sys = s_total.imag / 1000
    s_sys = abs(s_total) / 1000
    return round(p_sys, 5), round(q_sys, 5), round(s_sys, 5)
