"""
Testcase_AcuHMI_AcuRev4100_VoltInt_Thr_026
原始编号: TC_VoltInt_Thr_001 ~ TC_VoltInt_Thr_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (5, True),                                                # TC_VoltInt_Thr_001
    (10, True),                                               # TC_VoltInt_Thr_002
    (20, True),                                               # TC_VoltInt_Thr_003
    (4, False),                                               # TC_VoltInt_Thr_004
    (21, False),                                              # TC_VoltInt_Thr_005
]


def test_Testcase_AcuHMI_AcuRev4100_VoltInt_Thr_026(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltInt_Thr_026
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_threshold(app_page, INT_IDX, value)
        saved = save_and_check(app_page)
        _log.info("[TC] VoltIntrThreshold=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Voltage Interruption Threshold={value}: 保存失败"
            verify_modbus(REG_VOLT_INT_THR, value, label=f"VoltIntrThreshold({value})")
        else:
            assert not saved, f"Voltage Interruption Threshold={value}: 期望校验失败但保存成功"


