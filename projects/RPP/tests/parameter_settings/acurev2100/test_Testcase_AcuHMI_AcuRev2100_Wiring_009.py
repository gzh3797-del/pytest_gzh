"""
Testcase_AcuHMI_AcuRev2100_Wiring_009
原始编号: TC_Basic_006 ~ TC_Basic_011
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option ─────────────────────────────────────────
_CASES = [
    ('1 Element 2 Wire (1LN)'),                               # TC_Basic_006
    ('3 Element 4 Wire Y (3LN)'),                             # TC_Basic_007
    ('2 Element 3 Wire 1 Phase (2LN)'),                       # TC_Basic_008
    ('3 Element 4 Wire Delta (High-leg delta)'),              # TC_Basic_009
    ('2 Element 3 Wire Delta (2LL)'),                         # TC_Basic_010
    ('2 Element 3 Wire Network (2LN/Network)'),               # TC_Basic_011
]


def test_Testcase_AcuHMI_AcuRev2100_Wiring_009(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_Wiring_009
    参数组: option
    """
    for option in _CASES:
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "Wiring of Three Phase User", option)
        saved = save_and_check(app_page)
        assert saved, f"Wiring={option}: 保存失败"
        verify_modbus(REG_WIRING, WIRING_ENCODE[option], label=f"Wiring({option})")


