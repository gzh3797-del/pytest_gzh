"""
Testcase_AcuHMI_AcuvimIIW_010_SubInterval
原始编号: TC_SubInterval_001 ~ TC_SubInterval_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_SubInterval_001
    (15, True),                                               # TC_SubInterval_002
    (30, True),                                               # TC_SubInterval_003
    (0, False),                                               # TC_SubInterval_004
    (31, False),                                              # TC_SubInterval_005
]


def test_Testcase_AcuHMI_AcuvimIIW_010_SubInterval(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_010_SubInterval
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        _log.info("[TC] ════SubInterval=%s  ok=%s", value, ok)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_input(app_page, "Sub-Interval", value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"Sub-Interval={value}: 保存失败"
            _log.info("[EXPECT  ] SubInterval reg = %s", value)
            verify_reg(REG_DEMAND_SLIP, value, label=f"SubInterval({value})")
        else:
            assert not saved, f"Sub-Interval={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_SUB_INTERVAL)


