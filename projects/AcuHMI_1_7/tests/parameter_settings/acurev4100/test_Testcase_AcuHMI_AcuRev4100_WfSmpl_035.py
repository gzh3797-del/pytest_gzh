"""
Testcase_AcuHMI_AcuRev4100_WfSmpl_035
原始编号: TC_WfSmpl_001 ~ TC_WfSmpl_004
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('16 points', 0),                                         # TC_WfSmpl_001
    ('32 points', 1),                                         # TC_WfSmpl_002
    ('64 points', 2),                                         # TC_WfSmpl_003
    ('128 points', 3),                                        # TC_WfSmpl_004
]


def test_Testcase_AcuHMI_AcuRev4100_WfSmpl_035(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_WfSmpl_035
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_event_waveform(app_page)
        # 先将 Pre/Post Cycles 设为 1，避免与较大 Sample Rate 相乘超出 2560 限制
        set_input(app_page, "Num of Cycles Before", 1)
        set_input(app_page, "Num of Cycles After", 1)
        set_dropdown(app_page, "Sample Rate", option)
        saved = save_and_check(app_page)
        assert saved, f"Sampling Rate={option}: 保存失败"
        verify_modbus(REG_WF_SAMPLE_RATE, encoded, label=f"SamplingRate({option})")


