"""
Testcase_AcuHMI_AcuRev4100_VoltSag_RO_029
原始编号: TC_VoltSag_RO_001 ~ TC_VoltSag_RO_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('None', 0),                                              # TC_VoltSag_RO_001
    ('RO1', 1),                                               # TC_VoltSag_RO_002
    ('RO2', 2),                                               # TC_VoltSag_RO_003
]


def test_Testcase_AcuHMI_AcuRev4100_VoltSag_RO_029(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltSag_RO_029
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_ro(app_page, SAG_IDX, option)
        saved = save_and_check(app_page)
        assert saved, f"Volt Sag RO={option}: 保存失败"
        verify_modbus(REG_VOLT_SAG_RO, encoded, label=f"VoltSagRO({option})")


