"""
Testcase_AcuHMI_AcuvimIIW_AvgInterval_001
原始编号: TC_AvgInterval_001 ~ TC_AvgInterval_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_AvgInterval_001
    (15, True),                                               # TC_AvgInterval_002
    (30, True),                                               # TC_AvgInterval_003
    (0, False),                                               # TC_AvgInterval_004
    (31, False),                                              # TC_AvgInterval_005
]


def test_Testcase_AcuHMI_AcuvimIIW_AvgInterval_001(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_AvgInterval_001
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        _log.info("[TC] ════AvgInterval=%s  ok=%s", value, ok)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        if ok:
            # AvgInterval 必须 ≥ Sub-Interval；先将 Sub-Interval 重置为 1 避免交叉约束拒绝保存
            set_input(app_page, "Sub-Interval", 1)
        set_input(app_page, "Averaging Interval Window", value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"Averaging Interval Window={value}: 保存失败"
            _log.info("[EXPECT  ] AvgInterval reg = %s", value)
            verify_reg(REG_DEMAND_INTERVAL, value, label=f"AvgIntervalWindow({value})")
        else:
            assert not saved, f"Averaging Interval Window={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_AVG_INTERVAL)


