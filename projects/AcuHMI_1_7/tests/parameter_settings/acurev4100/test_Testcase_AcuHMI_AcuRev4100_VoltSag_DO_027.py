"""
Testcase_AcuHMI_AcuRev4100_VoltSag_DO_027
原始编号: TC_VoltSag_DO_001 ~ TC_VoltSag_DO_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('None', 0),                                              # TC_VoltSag_DO_001
    ('DO1', 1),                                               # TC_VoltSag_DO_002
    ('DO8', 8),                                               # TC_VoltSag_DO_003
]


def test_Testcase_AcuHMI_AcuRev4100_VoltSag_DO_027(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltSag_DO_027
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_do(app_page, SAG_IDX, option)
        saved = save_and_check(app_page)
        assert saved, f"Volt Sag DO={option}: 保存失败"
        verify_modbus(REG_VOLT_SAG_DO, encoded, label=f"VoltSagDO({option})")


