"""
Testcase_AcuHMI_AcuRev2100_Ct2Type_002
原始编号: TC_Basic_012 ~ TC_Basic_013
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('RCT', 120),                                             # TC_Basic_012
    ('333mV', 333),                                           # TC_Basic_013
]


def test_Testcase_AcuHMI_AcuRev2100_Ct2Type_002(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_Ct2Type_002
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "CT2", option)
        saved = save_and_check(app_page)
        assert saved, f"CT2={option}: 保存失败"
        verify_modbus(REG_CT_TYPE, encoded, label=f"CT2({option})")


