"""
Testcase_AcuHMI_AcuRev4100_WfPre_034
原始编号: TC_WfPre_001 ~ TC_WfPre_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_WfPre_001
    (50, True),                                               # TC_WfPre_002
    (158, True),                                              # TC_WfPre_003
    (0, False),                                               # TC_WfPre_004
    (160, False),                                             # TC_WfPre_005
]


def test_Testcase_AcuHMI_AcuRev4100_WfPre_034(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_WfPre_034
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_dropdown(app_page, "Sample Rate", "16 points")   # 扩大总容量上限
        set_input(app_page, "Num of Cycles After", 1)        # Post=1，最小化占用
        set_input(app_page, "Num of Cycles Before", value)
        saved = save_and_check(app_page)
        _log.info("[TC] WfPreCycles=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Num of Cycles Before={value}: 保存失败"
            verify_modbus(REG_WF_PRE_CYCLES, value, label=f"PreCycles({value})")
        else:
            assert not saved, f"Num of Cycles Before={value}: 期望校验失败但保存成功"


