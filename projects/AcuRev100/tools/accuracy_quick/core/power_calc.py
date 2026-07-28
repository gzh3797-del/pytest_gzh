"""功率计算模块（独立，逻辑来自 power_calculate.py，不依赖原工程）。

⚠️ 单位口径：返回 **W / var / VA**，与 AcuRev-100(RACG) 寄存器一致
   （2026-07 台面实证：220V × 50A、PF=1 → SYSTEM/PHASE_ACTIVE_POWER 读 11000，即 W）。
   1320 老版本这里除了 1000 返 kW，AcuRev-100 沿用会差 1000 倍。
"""
import math
import cmath


def active_power(voltage: float, current: float, v_angle_deg: float, i_angle_deg: float) -> float:
    """有功功率 P = U × I × cos(θ_V − θ_I)，单位 W"""
    angle = v_angle_deg - i_angle_deg
    return voltage * current * math.cos(math.radians(angle))


def reactive_power(voltage: float, current: float, v_angle_deg: float, i_angle_deg: float) -> float:
    """无功功率 Q = U × I × sin(θ_V − θ_I)，单位 var"""
    angle = v_angle_deg - i_angle_deg
    return voltage * current * math.sin(math.radians(angle))


def apparent_power(voltage: float, current: float) -> float:
    """视在功率 S = U × I，单位 VA"""
    return voltage * current


def power_factor(v_angle_deg: float, i_angle_deg: float) -> float:
    """功率因数 PF = cos(θ_V − θ_I)"""
    return math.cos(math.radians(v_angle_deg - i_angle_deg))


def phase_power(v: float, i: float, v_ang: float, i_ang: float) -> tuple[float, float, float]:
    """单相 (P, Q, S)，单位 W / var / VA"""
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
