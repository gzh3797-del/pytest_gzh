"""
Testcase_AcuHMI_AcuRev2100_VarPf_008
原始编号: TC_Advanced_003 ~ TC_Advanced_004
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('IEC', 0),                                               # TC_Advanced_003
    ('IEEE', 1),                                              # TC_Advanced_004
]


def test_Testcase_AcuHMI_AcuRev2100_VarPf_008(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_VarPf_008
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page, "Advanced")
        set_dropdown(app_page, "VAR/PF Convention", option)
        saved = save_and_check(app_page)
        assert saved, f"VAR/PF={option}: 保存失败"
        verify_modbus(REG_VAR_PF, encoded, label=f"VAR/PF({option})")


