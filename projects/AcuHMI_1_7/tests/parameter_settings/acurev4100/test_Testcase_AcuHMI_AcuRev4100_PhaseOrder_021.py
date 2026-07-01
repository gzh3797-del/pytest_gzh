"""
Testcase_AcuHMI_AcuRev4100_PhaseOrder_021
原始编号: TC_PhaseOrder_001 ~ TC_PhaseOrder_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('ABC', 0),                                               # TC_PhaseOrder_001
    ('ACB', 1),                                               # TC_PhaseOrder_002
]


def test_Testcase_AcuHMI_AcuRev4100_PhaseOrder_021(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_PhaseOrder_021
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Phase Order", option)
        saved = save_and_check(app_page)
        assert saved, f"Phase Order={option}: 保存失败"
        verify_modbus(REG_PHASE_ORDER, encoded, label=f"PhaseOrder({option})")


