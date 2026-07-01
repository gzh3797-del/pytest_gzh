"""
Testcase_AcuHMI_AcuRev4100_PT_022
原始编号: TC_PT_001 ~ TC_PT_012
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：pt1, pt2, ok, note ─────────────────────────────────────────
_CASES = [
    (51, 50, True, 'PT1 有效最小值'),                              # TC_PT_001
    (480, 50, True, 'PT1 典型值'),                               # TC_PT_002
    (1000000, 50, True, 'PT1 最大值'),                           # TC_PT_003
    (49, 50, False, 'PT1 低于下限'),                              # TC_PT_004
    (1000001, 50, False, 'PT1 超出上限'),                         # TC_PT_005
    (1000, 50, True, 'PT2 最小值'),                              # TC_PT_006
    (1000, 830, True, 'PT2 最大值'),                             # TC_PT_007
    (1000, 49, False, 'PT2 低于下限'),                            # TC_PT_008
    (1000, 831, False, 'PT2 超出上限'),                           # TC_PT_009
    (500, 499, True, 'PT2=PT1-1 刚好满足'),                       # TC_PT_010
    (500, 500, True,  'PT2=PT1 固件允许保存成功'),                   # TC_PT_011
    (500, 501, False, 'PT2>PT1 不满足约束'),                       # TC_PT_012
]


def test_Testcase_AcuHMI_AcuRev4100_PT_022(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_PT_022
    参数组: pt1, pt2, ok, note
    """
    for (pt1, pt2, ok, note) in _CASES:
        nav_to_general(app_page)
        # 先将 Nominal Voltage 设为 50（≤ 最小 PT1=51），避免"Nominal Voltage must be ≤ PT1"拦截
        set_input(app_page, "Nominal Voltage", 50)
        set_input(app_page, "PT2", pt2)
        set_input(app_page, "PT1", pt1)
        saved = save_and_check(app_page)
        _log.info("[TC] PT1=%-8s PT2=%-5s ok=%s  note=%s  result=%s",
                  pt1, pt2, ok, note, saved)
        if ok:
            assert saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望成功但出现错误"
            verify_modbus_32(REG_PT1, pt1, label="PT1 Primary Voltage")
            verify_modbus(REG_PT2, pt2, label="PT2 Secondary Voltage")
        else:
            assert not saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望校验失败但保存成功"


