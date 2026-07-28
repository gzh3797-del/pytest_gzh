"""
Testcase_AcuHMI_AcuRev4100_Demand_004
原始编号: TC_Demand_001 ~ TC_Demand_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Fixed', 0),                                             # TC_Demand_001
    ('Sliding', 1),                                           # TC_Demand_002
]


def test_Testcase_AcuHMI_AcuRev4100_Demand_004(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_Demand_004
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Demand Method", option)
        saved = save_and_check(app_page)
        assert saved, f"Demand Method={option}: 保存失败"
        verify_modbus(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


