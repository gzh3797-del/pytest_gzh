"""
Testcase_AcuHMI_AcuRev2100_AveragingInterval_001
原始编号: TC_Basic_026 ~ TC_Basic_030
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok, err_msg ─────────────────────────────────────────
_CASES = [
    (2, True, None),                                          # TC_Basic_026
    (15, True, None),                                         # TC_Basic_027
    (30, True, None),                                         # TC_Basic_028
    (0, False, _ERR_AVG_INTERVAL),                            # TC_Basic_029
    (31, False, _ERR_AVG_INTERVAL),                           # TC_Basic_030
]


def test_Testcase_AcuHMI_AcuRev2100_AveragingInterval_001(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_AveragingInterval_001
    参数组: value, ok, err_msg
    """
    for (value, ok, err_msg) in _CASES:
        step("前置：设 Sub-Interval = 1")
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "Sliding Window", "Rolling Window Demand")
        app_page.wait_for_timeout(500)
        set_input(app_page, "Sub-Interval", 1)
        save_and_check(app_page)

        step(f"测试：设 Averaging Interval Window = {value}")
        nav_to_general(app_page, "Basic")
        set_input(app_page, "Averaging Interval Window", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Averaging Interval=%s  ok_expected=%s  save_result=%s", value, ok, saved)
        if ok:
            assert saved, f"Averaging Interval={value}: 保存失败"
            verify_modbus(REG_DEMAND_INTERVAL, value, label="Averaging Interval Window")
        else:
            assert not saved, f"Averaging Interval={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, err_msg)


