"""
Testcase_AcuHMI_AcuvimIIW_003_CT1
原始编号: TC_CT1_001 ~ TC_CT1_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (5, True),                                                # TC_CT1_001
    (100, True),                                              # TC_CT1_002
    (50000, True),                                            # TC_CT1_003
    (0, False),                                               # TC_CT1_004
    (50001, False),                                           # TC_CT1_005
]


def test_Testcase_AcuHMI_AcuvimIIW_003_CT1(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_003_CT1
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        _log.info("[TC] ════CT1=%s  ok=%s", value, ok)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_input(app_page, "CT1", value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"CT1={value}: 保存失败"
            _log.info("[EXPECT  ] CT1 reg = %s", value)
            verify_reg(REG_CT1, value, label=f"CT1({value})")
        else:
            assert not saved, f"CT1={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_CT1)


