import math
import cmath

# def line_to_line_voltage_calculate(ua, ub, uc, ua_p, ub_p, uc_p):
#     """
#     计算线电压（幅值和复数形式）
#     :param ua, ub, uc: 相电压幅值
#     :param ua_p, ub_p, uc_p: 相电压相角，单位°
#     :return: (vab_abs, vbc_abs, vca_abs), (vab_c, vbc_c, vca_c)
#     """
#     va = cmath.rect(ua, math.radians(ua_p))
#     vb = cmath.rect(ub, math.radians(ub_p))
#     vc = cmath.rect(uc, math.radians(uc_p))
#
#     vab_c = va - vb
#     vbc_c = vb - vc
#     vca_c = vc - va
#
#     vab_abs = abs(vab_c)
#     vbc_abs = abs(vbc_c)
#     vca_abs = abs(vca_c)
#
#     return (vab_abs, vbc_abs, vca_abs), (vab_c, vbc_c, vca_c)


def calculate_2e3wd_power(ua, ub, uc, ua_p, ub_p, uc_p, ia, ib, ic, ia_p, ib_p, ic_p):
    """
    2e3wD sys功率计算方法
    Args:
        ua: 相电压A
        ub: 相电压B
        uc: 相电压C
        ua_p: 相电压A的相角
        ub_p: 相电压B的相角
        uc_p: 相电压C的相角
        ia: 电流A
        ib: 电流B
        ic: 电流C
        ia_p: 电流A的相角
        ib_p: 电流B的相角
        ic_p: 电流C的相角
    计算 U_AB/I_A 和 U_CB/I_C 相位差，并打印中间角度
    :return: p_sys, q_sys, s_sys
    """
    # 构建复数相量
    va = cmath.rect(ua, math.radians(ua_p))
    vb = cmath.rect(ub, math.radians(ub_p))
    vc = cmath.rect(uc, math.radians(uc_p))

    ia_c = cmath.rect(ia, math.radians(ia_p))
    ic_c = cmath.rect(ic, math.radians(ic_p))

    vab = va - vb
    vcb = vc - vb

    # 打印中间相角
    print(f"U_AB 相角: {math.degrees(cmath.phase(vab)):.2f}°")
    print(f"U_CB 相角: {math.degrees(cmath.phase(vcb)):.2f}°")
    print(f"I_A 相角: {math.degrees(cmath.phase(ia_c)):.2f}°")
    print(f"I_C 相角: {math.degrees(cmath.phase(ic_c)):.2f}°")

    # 相位差
    ua_ia_angle = math.degrees(cmath.phase(vab / ia_c))
    uc_ic_angle = math.degrees(cmath.phase(vcb / ic_c))

    # 归一化到 [-180, 180]
    ua_ia_angle = (ua_ia_angle + 180) % 360 - 180
    uc_ic_angle = (uc_ic_angle + 180) % 360 - 180
    # 打印中间相角
    print(f"U_AB与Ia夹角: {ua_ia_angle:.2f}°")
    print(f"U_CB与Ic夹角: {uc_ic_angle:.2f}°")

    # 有功、无功功率（复数法，考虑两条线）
    s_total = vab * ia_c.conjugate() + vcb * ic_c.conjugate()

    p_sys = s_total.real
    q_sys = s_total.imag
    s_sys = abs(s_total)

    return p_sys, q_sys, s_sys



# ================== 使用示例 ==================


ua, ub, uc = 100, 100, 100
ua_p, ub_p, uc_p = 0, 240, 120
ia, ib, ic = 2, 0, 2
ia_p, ib_p, ic_p = 30, 270, 150

# 有功、无功功率（复数法，考虑两条线）

p_sys, q_sys, s_sys = calculate_2e3wd_power(ua, ub, uc, ua_p, ub_p, uc_p, ia, ib, ic, ia_p, ib_p, ic_p)

print(f"P_sys: {p_sys:.3f} W")
print(f"Q_sys: {q_sys:.3f} var")
print(f"S_sys: {s_sys:.3f} VA")
