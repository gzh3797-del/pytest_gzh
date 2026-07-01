"""
接线检查算法引擎 —— 根据控源输出参数推导预期 Wiring Status。

依据：接线检测总表_ver1.05.xlsx 前5个 Sheet 的算法规格。
实现：3E4WY / 2E3W Delta / 2E3W Network / 2E3W 1Phase / 1E2W。

输出结构：
{
    'voltage': {
        'A': str,      # Phase A 电压状态
        'B': str,      # Phase B 电压状态
        'C': str,      # Phase C 电压状态
        'order': str,  # Phase Order 状态
    },
    'current': {
        'A': str,   # Ia 电流状态
        'B': str,   # Ib 电流状态
        'C': str,   # Ic 电流状态
    }
}
状态字符串与 HMI 页面 Wiring Status 列文字一致。
"""
import cmath
import math

# ── 工具函数 ─────────────────────────────────────────────────────────────────

def _phasor(amp: float, deg: float) -> complex:
    return cmath.rect(amp, math.radians(deg))


def _line_voltage(amp_a, deg_a, amp_b, deg_b) -> float:
    """两相电压相量差的幅值（线电压）"""
    return abs(_phasor(amp_a, deg_a) - _phasor(amp_b, deg_b))


def _angle_in_range(angle: float, center: float, tolerance: float) -> bool:
    """判断角度是否在 center ± tolerance 内（考虑360°回绕）"""
    diff = abs(((angle - center + 180) % 360) - 180)
    return diff <= tolerance


def _pf(v_deg: float, i_deg: float) -> float:
    """计算功率因数 cos(∠V - ∠I)"""
    return math.cos(math.radians((v_deg - i_deg) % 360))


def _v_unbalance(va: float, vb: float, vc: float) -> float:
    """电压不平衡度（%）= max(|Vi - Vavg|) / Vavg × 100"""
    avg = (va + vb + vc) / 3
    if avg == 0:
        return float('inf')
    return max(abs(va - avg), abs(vb - avg), abs(vc - avg)) / avg * 100


# ── 电流单相状态 ──────────────────────────────────────────────────────────────

def _current_status(i_rms: float, pf_val: float) -> str:
    if i_rms < 0.1:
        return 'Wiring Missing'
    if -1.0 <= pf_val <= -0.9:
        return 'Polarity Reversed'
    if -0.9 < pf_val <= 0.9:
        return 'Phase Shift'
    return 'Pass'


def _current_status_v(i_rms: float, v_amp: float, v_deg: float,
                      i_deg: float, vrate: float) -> str:
    """
    带电压幅值检查的电流状态计算。
    当相电压幅值 < 0.1·VRATE 时，设备无有效相角参考，PF 无意义；
    与 2E3W 1Phase 引擎保持一致：有电流 → Phase Shift，无电流 → Wiring Missing。
    """
    if v_amp < 0.1 * vrate:
        return 'Wiring Missing' if i_rms < 0.1 else 'Phase Shift'
    return _current_status(i_rms, _pf(v_deg, i_deg))


# ── 3E4WY 主算法 ─────────────────────────────────────────────────────────────

def compute_3e4wy(
    ua: float, qua: float,
    ub: float, qub: float,
    uc: float, quc: float,
    ia: float, qia: float,
    ib: float, qib: float,
    ic: float, qic: float,
    vrate: float,
    phase_order: int,          # 0=ABC, 1=ACB
) -> dict:
    """
    计算 3E4WY 接线检查预期结果。

    参数：
        ua/qua/ub/qub/uc/quc  — A/B/C 相电压幅值(V)和相角(°)
        ia/qia/ib/qib/ic/qic  — A/B/C 相电流幅值(A)和相角(°)
        vrate                  — 额定电压(V)
        phase_order            — 相序：0=ABC, 1=ACB
    """
    Van, Vbn, Vcn = ua, ub, uc
    Vab = _line_voltage(ua, qua, ub, qub)
    Vbc = _line_voltage(ub, qub, uc, quc)
    Vca = _line_voltage(uc, quc, ua, qua)

    va_s = vb_s = vc_s = 'Pass'
    order_s = 'Pass'

    # ── Step 1：接线缺失（按顺序，满足则停止）───────────────────────────────
    if Van < 0.1*vrate and Vbn < 0.1*vrate and Vcn < 0.1*vrate:      # 条件1
        va_s = 'Va Wiring Missing'
        vb_s = 'Vb Wiring Missing'
        vc_s = 'Vc Wiring Missing'
    elif Vab < 0.1*vrate and 0.8*vrate <= Vcn <= 1.2*vrate:           # 条件2
        va_s = 'Va Wiring Missing'
        vb_s = 'Vb Wiring Missing'
    elif Vbc < 0.1*vrate and 0.8*vrate <= Van <= 1.2*vrate:           # 条件3
        vb_s = 'Vb Wiring Missing'
        vc_s = 'Vc Wiring Missing'
    elif Vca < 0.1*vrate and 0.8*vrate <= Vbn <= 1.2*vrate:           # 条件4
        va_s = 'Va Wiring Missing'
        vc_s = 'Vc Wiring Missing'
    elif Van <= 0.8*vrate:                                             # 条件5
        va_s = 'Va Wiring Missing'
    elif Vbn <= 0.8*vrate:                                             # 条件6
        vb_s = 'Vb Wiring Missing'
    elif Vcn <= 0.8*vrate:                                             # 条件7
        vc_s = 'Vc Wiring Missing'

    # ── Step 2：反接（无缺失时执行，按顺序，满足则停止）──────────────────────
    elif 0.8*vrate < Van < 1.2*vrate and Vbn > 1.3*vrate and Vcn > 1.3*vrate:   # 条件8
        va_s = 'Va-Vn Reversed'
    elif 0.8*vrate < Vbn < 1.2*vrate and Van > 1.3*vrate and Vcn > 1.3*vrate:   # 条件9
        vb_s = 'Vb-Vn Reversed'
    elif 0.8*vrate < Vcn < 1.2*vrate and Van > 1.3*vrate and Vbn > 1.3*vrate:   # 条件10
        vc_s = 'Vc-Vn Reversed'

    else:
        # ── Step 3：相位错误（无反接时执行，条件11-13 相互独立）──────────────
        if phase_order == 0:   # ABC
            if not _angle_in_range(qub, 240, 20):   # 条件11
                vb_s = 'Phase Shift'
            if not _angle_in_range(quc, 120, 20):   # 条件12
                vc_s = 'Phase Shift'
        else:                  # ACB
            if not _angle_in_range(qub, 120, 20):   # 条件11
                vb_s = 'Phase Shift'
            if not _angle_in_range(quc, 240, 20):   # 条件12
                vc_s = 'Phase Shift'

        # 条件13：角度法检测相序错误
        if phase_order == 0:   # ABC 配置
            if _angle_in_range(qub, 120, 20) and _angle_in_range(quc, 240, 20):
                order_s = 'Phase Order Error'
        else:                  # ACB 配置
            if _angle_in_range(qub, 240, 20) and _angle_in_range(quc, 120, 20):
                order_s = 'Phase Order Error'

    # ── 电流侧（各相独立，与电压侧无关）─────────────────────────────────────
    ia_s = _current_status_v(ia, ua, qua, qia, vrate)
    ib_s = _current_status_v(ib, ub, qub, qib, vrate)
    ic_s = _current_status_v(ic, uc, quc, qic, vrate)

    return {
        'voltage': {'A': va_s, 'B': vb_s, 'C': vc_s, 'order': order_s},
        'current': {'A': ia_s, 'B': ib_s, 'C': ic_s},
    }


# ── 2E3W Network ──────────────────────────────────────────────────────────────
# 电压/电流算法与 3E4WY 完全相同
compute_2e3w_network = compute_3e4wy


# ── 2E3W Delta ────────────────────────────────────────────────────────────────

def _phasor_angle(amp_a, deg_a, amp_b, deg_b) -> float:
    """两相相量差的角度（度）"""
    v = _phasor(amp_a, deg_a) - _phasor(amp_b, deg_b)
    return math.degrees(cmath.phase(v)) % 360


def _current_status_delta(i_rms: float, angle: float,
                           phase: str, phase_order: int) -> str:
    """2E3W Delta 电流状态（按角度判断）"""
    if i_rms < 0.1:
        return 'Wiring Missing'
    if phase == 'A':
        rev_center  = 150 if phase_order == 0 else 210
        pass_center = 330 if phase_order == 0 else  30
    else:  # C
        rev_center  = 270 if phase_order == 0 else  90
        pass_center =  90 if phase_order == 0 else 270
    if _angle_in_range(angle, rev_center, 20):
        return 'Polarity Reversed'
    if not _angle_in_range(angle, pass_center, 20):
        return 'Phase Shift'
    return 'Pass'


def compute_2e3w_delta(
    ua: float, qua: float,
    ub: float, qub: float,
    uc: float, quc: float,
    ia: float, qia: float,
    ic: float, qic: float,
    vrate: float,
    phase_order: int,
) -> dict:
    Vab = _line_voltage(ua, qua, ub, qub)
    Vbc = _line_voltage(ub, qub, uc, quc)
    Vca = _line_voltage(uc, quc, ua, qua)
    ang_Vab = _phasor_angle(ua, qua, ub, qub)
    ang_Vbc = (_phasor_angle(ub, qub, uc, quc) - ang_Vab) % 360
    ang_Vca = (_phasor_angle(uc, quc, ua, qua) - ang_Vab) % 360

    va_s = vb_s = vc_s = 'Pass'
    order_s = 'Pass'

    # Step 1：接线缺失
    if Vab < 0.1*vrate and Vbc < 0.1*vrate and Vca < 0.1*vrate:
        va_s = 'Va Wiring Missing'
        vb_s = 'Vb Wiring Missing'
        vc_s = 'Vc Wiring Missing'
    elif Vab < 0.1*vrate:
        va_s, vb_s = 'Va Wiring Missing', 'Vb Wiring Missing'
    elif Vbc < 0.1*vrate:
        vc_s, vb_s = 'Vc Wiring Missing', 'Vb Wiring Missing'
    elif Vca < 0.1*vrate:
        vc_s, va_s = 'Vc Wiring Missing', 'Va Wiring Missing'
    elif 0.8*vrate < Vbc < 1.2*vrate and (Vab < 0.8*vrate or Vca < 0.8*vrate):
        va_s = 'Va Wiring Missing'
    elif 0.8*vrate < Vca < 1.2*vrate and (Vab < 0.8*vrate or Vbc < 0.8*vrate):
        vb_s = 'Vb Wiring Missing'
    elif 0.8*vrate < Vab < 1.2*vrate and (Vbc < 0.8*vrate or Vca < 0.8*vrate):
        vc_s = 'Vc Wiring Missing'
    else:
        # Step 2：相序错误
        if phase_order == 0:   # ABC 配置：ACB 信号 → 错误
            if _angle_in_range(ang_Vbc, 120, 20) and _angle_in_range(ang_Vca, 240, 20):
                order_s = 'Phase Order Error'
        else:                  # ACB 配置：ABC 信号 → 错误
            if _angle_in_range(ang_Vbc, 240, 20) and _angle_in_range(ang_Vca, 120, 20):
                order_s = 'Phase Order Error'

    # 设备以 Vab 为基准测量电流相角；将输入绝对角度归一化为相对 Vab 的角度
    eff_qia = (qia - ang_Vab) % 360
    eff_qic = (qic - ang_Vab) % 360
    ia_s = _current_status_delta(ia, eff_qia, 'A', phase_order)
    ic_s = _current_status_delta(ic, eff_qic, 'C', phase_order)

    return {
        'voltage': {'A': va_s, 'B': vb_s, 'C': vc_s, 'order': order_s},
        'current': {'A': ia_s, 'B': 'N/A', 'C': ic_s},
    }


# ── 2E3W 1Phase ───────────────────────────────────────────────────────────────

def compute_2e3w_1phase(
    ua: float, qua: float,
    uc: float, quc: float,
    ia: float, qia: float,
    ic: float, qic: float,
    vrate: float,
) -> dict:
    Van, Vcn = ua, uc
    Vca = _line_voltage(uc, quc, ua, qua)

    va_s = vc_s = 'Pass'

    # Step 1：接线缺失
    if Vca < 0.1*vrate:
        va_s = 'Va Wiring Missing'
        vc_s = 'Vc Wiring Missing'
    elif Van < 0.8*vrate:
        va_s = 'Va Wiring Missing'
    elif Vcn < 0.8*vrate:
        vc_s = 'Vc Wiring Missing'
    # Step 2：反接
    elif 0.8*vrate < Van < 1.2*vrate and Vcn > 1.3*vrate:
        va_s = 'Va-Vn Reversed'
    elif 0.8*vrate < Vcn < 1.2*vrate and Van > 1.3*vrate:
        vc_s = 'Vc-Vn Reversed'

    # 电压幅值 < 0.1·VRATE 时无有效参考，PF 未定义：有电流→ Phase Shift，无电流→ Wiring Missing
    if ua < 0.1 * vrate:
        ia_s = 'Wiring Missing' if ia < 0.1 else 'Phase Shift'
    else:
        ia_s = _current_status(ia, _pf(qua, qia))

    if uc < 0.1 * vrate:
        ic_s = 'Wiring Missing' if ic < 0.1 else 'Phase Shift'
    else:
        ic_s = _current_status(ic, _pf(quc, qic))

    return {
        'voltage': {'A': va_s, 'B': 'N/A', 'C': vc_s, 'order': 'N/A'},
        'current': {'A': ia_s, 'B': 'N/A', 'C': ic_s},
    }


# ── 1E2W ──────────────────────────────────────────────────────────────────────

def compute_1e2w(
    ua: float, qua: float,
    ia: float, qia: float,
    vrate: float,
) -> dict:
    # 1E2W 规格只有三条：Missing / Polarity Reversed / Pass，无 Phase Shift
    va_s = 'Va Wiring Missing' if ua < 0.1*vrate else 'Pass'
    if ia < 0.1:
        ia_s = 'Wiring Missing'
    elif -1.0 <= _pf(qua, qia) <= -0.9:
        ia_s = 'Polarity Reversed'
    else:
        ia_s = 'Pass'

    return {
        'voltage': {'A': va_s, 'B': 'N/A', 'C': 'N/A', 'order': 'N/A'},
        'current': {'A': ia_s, 'B': 'N/A', 'C': 'N/A'},
    }
