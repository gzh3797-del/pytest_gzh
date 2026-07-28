"""
Testcase_AcuHMI_AcuRev4100_Freq_013
原始编号: TC_Freq_001 ~ TC_Freq_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('50Hz', 0),                                              # TC_Freq_001
    ('60Hz', 1),                                              # TC_Freq_002
]


def test_Testcase_AcuHMI_AcuRev4100_Freq_013(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_Freq_013
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Nominal Frequency", option)
        saved = save_and_check(app_page)
        assert saved, f"Nominal Frequency={option}: 保存失败"
        verify_modbus(REG_FREQ_SEL, encoded, label=f"FrequencySelection({option})")


