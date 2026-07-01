"""
Testcase_AcuHMI_AcuRev4100_VARPF_024
原始编号: TC_VARPF_001 ~ TC_VARPF_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('IEC', 0),                                               # TC_VARPF_001
    ('IEEE', 1),                                              # TC_VARPF_002
]


def test_Testcase_AcuHMI_AcuRev4100_VARPF_024(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VARPF_024
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "VAR/PF Convention", option)
        saved = save_and_check(app_page)
        assert saved, f"VAR/PF={option}: 保存失败"
        verify_modbus(REG_VAR_PF, encoded, label=f"VAR/PF({option})")


