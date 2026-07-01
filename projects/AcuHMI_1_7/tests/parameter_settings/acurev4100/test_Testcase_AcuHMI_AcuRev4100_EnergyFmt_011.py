"""
Testcase_AcuHMI_AcuRev4100_EnergyFmt_011
原始编号: TC_EnergyFmt_001 ~ TC_EnergyFmt_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('0.1 kWh', 0),                                           # TC_EnergyFmt_001
    ('0.01 kWh', 1),                                          # TC_EnergyFmt_002
    ('0.001 kWh', 2),                                         # TC_EnergyFmt_003
]


def test_Testcase_AcuHMI_AcuRev4100_EnergyFmt_011(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_EnergyFmt_011
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Energy Reading Display Format", option)
        saved = save_and_check(app_page)
        assert saved, f"Energy Reading Display Format={option}: 保存失败"
        verify_modbus(REG_ENERGY_FMT, encoded, label=f"EnergyFmt({option})")


