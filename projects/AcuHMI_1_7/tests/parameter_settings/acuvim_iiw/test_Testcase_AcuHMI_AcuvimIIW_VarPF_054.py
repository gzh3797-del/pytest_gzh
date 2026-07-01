"""
Testcase_AcuHMI_AcuvimIIW_VarPF_054
原始编号: TC_VarPF_001 ~ TC_VarPF_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('IEC', 0),                                               # TC_VarPF_001
    ('IEEE', 1),                                              # TC_VarPF_002
]


def test_Testcase_AcuHMI_AcuvimIIW_VarPF_054(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_VarPF_054
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════varPF=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Advanced")
        set_dropdown(app_page, "var/PF Convention", option)
        saved = save_and_check(app_page)
        assert saved, f"var/PF={option}: 保存失败"
        _log.info("[EXPECT  ] varPF reg = %s", encoded)
        verify_reg(REG_VAR_PF, encoded, label=f"varPF({option})")


