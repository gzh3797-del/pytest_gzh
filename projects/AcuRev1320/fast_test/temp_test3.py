import math
import cmath
import time

def calculate_2e3wn_power(ua, ub, uc, ua_p, ub_p, uc_p,
                          ia, ib, ic, ia_p, ib_p, ic_p):
    """
    2E3WN 系统功率计算（P_sys, Q_sys, S_sys）
    :param ua, ub, uc: 相电压幅值
    :param ua_p, ub_p, uc_p: 相电压相角，单位°
    :param ia, ib, ic: 相电流幅值
    :param ia_p, ib_p, ic_p: 相电流相角，单位°
    :return: P_sys, Q_sys, S_sys
    """
    # 构建复数相量
    va = cmath.rect(ua, math.radians(ua_p))
    vb = cmath.rect(ub, math.radians(ub_p))
    vc = cmath.rect(uc, math.radians(uc_p))

    ia_c = cmath.rect(ia, math.radians(ia_p))
    ic_c = cmath.rect(ic, math.radians(ic_p))

    # 线电压复数
    vab = va - vb
    vcb = vc - vb

    # 系统复功率（Aron 两元件法）
    s_total = vab * ia_c.conjugate() + vcb * ic_c.conjugate()

    p_sys = s_total.real
    q_sys = s_total.imag
    s_sys = abs(s_total)

    return p_sys, q_sys, s_sys

ua, ub, uc = 100, 100, 100
ua_p, ub_p, uc_p = 0, 240, 120
ia, ib, ic = 2, 2, 2
ia_p, ib_p, ic_p = 30, 270, 150

P, Q, S = calculate_2e3wn_power(
    ua, ub, uc, ua_p, ub_p, uc_p,
    ia, ib, ic, ia_p, ib_p, ic_p
)

print(f"P_sys = {P:.2f} W")
print(f"Q_sys = {Q:.2f} var")
print(f"S_sys = {S:.2f} VA")

print(f"======================{time.strftime('%Y_%m_%d %H:%M:%S')}======================")
start_time = time.time()
time.sleep(30)
print(f"====================测试总耗时:{time.time() - start_time}====================")
print(f"====================={time.strftime('%Y_%m_%d %H:%M:%S')}=====================")