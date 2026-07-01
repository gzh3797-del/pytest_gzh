"""
Testcase_AcuHMI_AcuRev4100_VoltSwl_Thr_032
原始编号: TC_VoltSwl_Thr_001 ~ TC_VoltSwl_Thr_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (110, True),                                              # TC_VoltSwl_Thr_001
    (130, True),                                              # TC_VoltSwl_Thr_002
    (150, True),                                              # TC_VoltSwl_Thr_003
    (109, False),                                             # TC_VoltSwl_Thr_004
    (151, False),                                             # TC_VoltSwl_Thr_005
]


def test_Testcase_AcuHMI_AcuRev4100_VoltSwl_Thr_032(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltSwl_Thr_032
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_threshold(app_page, SWL_IDX, value)
        saved = save_and_check(app_page)
        _log.info("[TC] VoltSwellThreshold=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Voltage Swell Threshold={value}: 保存失败"
            verify_modbus(REG_VOLT_SWL_THR, value, label=f"VoltSwellThreshold({value})")
        else:
            assert not saved, f"Voltage Swell Threshold={value}: 期望校验失败但保存成功"


