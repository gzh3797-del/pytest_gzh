"""
Testcase_AcuHMI_AcuRev4100_VoltSag_Thr_030
原始编号: TC_VoltSag_Thr_001 ~ TC_VoltSag_Thr_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (10, True),                                               # TC_VoltSag_Thr_001
    (50, True),                                               # TC_VoltSag_Thr_002
    (90, True),                                               # TC_VoltSag_Thr_003
    (9, False),                                               # TC_VoltSag_Thr_004
    (91, False),                                              # TC_VoltSag_Thr_005
]


def test_Testcase_AcuHMI_AcuRev4100_VoltSag_Thr_030(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltSag_Thr_030
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_threshold(app_page, SAG_IDX, value)
        saved = save_and_check(app_page)
        _log.info("[TC] VoltSagThreshold=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Voltage Sag Threshold={value}: 保存失败"
            verify_modbus(REG_VOLT_SAG_THR, value, label=f"VoltSagThreshold({value})")
        else:
            assert not saved, f"Voltage Sag Threshold={value}: 期望校验失败但保存成功"


