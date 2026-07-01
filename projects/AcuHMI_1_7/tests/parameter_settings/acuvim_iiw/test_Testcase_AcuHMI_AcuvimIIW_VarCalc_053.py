"""
Testcase_AcuHMI_AcuvimIIW_VarCalc_053
原始编号: TC_VarCalc_001 ~ TC_VarCalc_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('True Reactive Power', 0),                               # TC_VarCalc_001
    ('Generic Reactive Power', 1),                            # TC_VarCalc_002
]


def test_Testcase_AcuHMI_AcuvimIIW_VarCalc_053(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_VarCalc_053
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════varCalcMethod=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Advanced")
        set_dropdown(app_page, "var Calculation Method", option)
        saved = save_and_check(app_page)
        assert saved, f"var Calc Method={option}: 保存失败"
        _log.info("[EXPECT  ] varCalcMethod reg = %s", encoded)
        verify_reg(REG_REACTIVE_METHOD, encoded, label=f"varCalcMethod({option})")


