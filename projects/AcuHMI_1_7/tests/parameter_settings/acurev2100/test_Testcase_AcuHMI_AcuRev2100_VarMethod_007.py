"""
Testcase_AcuHMI_AcuRev2100_VarMethod_007
原始编号: TC_Advanced_001 ~ TC_Advanced_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Method 1 (True)', 0),                                   # TC_Advanced_001
    ('Method 2 (Generalized)', 1),                            # TC_Advanced_002
]


def test_Testcase_AcuHMI_AcuRev2100_VarMethod_007(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_VarMethod_007
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page, "Advanced")
        set_dropdown(app_page, "VAR Calculation Method", option)
        saved = save_and_check(app_page)
        assert saved, f"VAR Method={option}: 保存失败"
        verify_modbus(REG_VAR_METHOD, encoded, label=f"VAR Method({option})")


