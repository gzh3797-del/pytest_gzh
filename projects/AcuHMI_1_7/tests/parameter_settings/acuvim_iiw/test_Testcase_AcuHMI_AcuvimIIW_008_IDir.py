"""
Testcase_AcuHMI_AcuvimIIW_008_IDir
原始编号: TC_IDir_A_Pos ~ TC_IDir_C_Neg
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：field, reg, option, encoded ─────────────────────────────────────────
_CASES = [
    ('I A Direction', REG_IA_DIR, 'Positive', 0),             # TC_IDir_A_Pos
    ('I A Direction', REG_IA_DIR, 'Negative', 1),             # TC_IDir_A_Neg
    ('I B Direction', REG_IB_DIR, 'Positive', 0),             # TC_IDir_B_Pos
    ('I B Direction', REG_IB_DIR, 'Negative', 1),             # TC_IDir_B_Neg
    ('I C Direction', REG_IC_DIR, 'Positive', 0),             # TC_IDir_C_Pos
    ('I C Direction', REG_IC_DIR, 'Negative', 1),             # TC_IDir_C_Neg
]


def test_Testcase_AcuHMI_AcuvimIIW_008_IDir(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_008_IDir
    参数组: field, reg, option, encoded
    """
    for (field, reg, option, encoded) in _CASES:
        _log.info("[TC] ════field=%s  option=%s  expect_reg=%s", field, option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_radio(app_page, field, option)
        saved = save_and_check(app_page)
        assert saved, f"{field}={option}: 保存失败"
        _log.info("[EXPECT  ] %s reg = %s", field, encoded)
        verify_reg(reg, encoded, label=f"{field}({option})")


