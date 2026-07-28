"""
Testcase_AcuHMI_AcuRev4100_ManTrig_016
原始编号: TC_ManTrig_001 ~ TC_ManTrig_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Disable', 0),                                           # TC_ManTrig_001
    ('User 1', 1),                                            # TC_ManTrig_002
    ('Only Record Voltage', 255),                             # TC_ManTrig_003
]


def test_Testcase_AcuHMI_AcuRev4100_ManTrig_016(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_ManTrig_016
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_event_waveform(app_page)
        set_dropdown(app_page, "Manually Trigger", option)
        saved = save_and_check(app_page)
        assert saved, f"Manually Trigger={option}: 保存失败"
        verify_modbus(REG_WF_MAN_TRIG, encoded, label=f"ManuallyTrigger({option})")


