"""
Testcase_AcuHMI_AcuRev2100_SubInterval_006
原始编号: TC_Basic_021 ~ TC_Basic_025
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok, err_msg, avg_pre ─────────────────────────────────────────
_CASES = [
    (1, True, None, 30),                                      # TC_Basic_021
    (15, True, None, 30),                                     # TC_Basic_022
    (15, True, None, 30),                                     # TC_Basic_023
    (0, False, _ERR_SUB_INTERVAL, 30),                        # TC_Basic_024
    (31, False, _ERR_SUB_INTERVAL, 30),                       # TC_Basic_025
]


def test_Testcase_AcuHMI_AcuRev2100_SubInterval_006(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_SubInterval_006
    参数组: value, ok, err_msg, avg_pre
    """
    for (value, ok, err_msg, avg_pre) in _CASES:
        step(f"前置：设 Averaging Interval Window = {avg_pre}")
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
        app_page.wait_for_timeout(500)
        set_input(app_page, "Averaging Interval Window", avg_pre)
        save_and_check(app_page)

        step(f"测试：设 Sub-Interval = {value}")
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
        app_page.wait_for_timeout(500)
        set_input(app_page, "Sub-Interval", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Sub-Interval=%s  ok_expected=%s  save_result=%s", value, ok, saved)
        if ok:
            assert saved, f"Sub-Interval={value}: 保存失败"
            verify_modbus(REG_DEMAND_SUB, value, label="Sub-Interval")
        else:
            assert not saved, f"Sub-Interval={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, err_msg)


