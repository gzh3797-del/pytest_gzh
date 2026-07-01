"""
Testcase_AcuHMI_AcuRev4100_WfPost_033
原始编号: TC_WfPost_001 ~ TC_WfPost_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_WfPost_001
    (50, True),                                               # TC_WfPost_002
    (158, True),                                              # TC_WfPost_003
    (0, False),                                               # TC_WfPost_004
    (160, False),                                             # TC_WfPost_005
]


def test_Testcase_AcuHMI_AcuRev4100_WfPost_033(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_WfPost_033
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_dropdown(app_page, "Sample Rate", "16 points")
        set_input(app_page, "Num of Cycles Before", 1)       # Pre=1，最小化占用
        set_input(app_page, "Num of Cycles After", value)
        saved = save_and_check(app_page)
        _log.info("[TC] WfPostCycles=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Num of Cycles After={value}: 保存失败"
            verify_modbus(REG_WF_POST_CYCLES, value, label=f"PostCycles({value})")
        else:
            assert not saved, f"Num of Cycles After={value}: 期望校验失败但保存成功"


