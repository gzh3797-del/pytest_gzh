"""
Testcase_AcuHMI_AcuRev4100_VoltSwl_Hys_031
原始编号: TC_VoltSwl_Hys_001 ~ TC_VoltSwl_Hys_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_VoltSwl_Hys_001
    (5, True),                                                # TC_VoltSwl_Hys_002
    (10, True),                                               # TC_VoltSwl_Hys_003
    (0, False),                                               # TC_VoltSwl_Hys_004
    (11, False),                                              # TC_VoltSwl_Hys_005
]


def test_Testcase_AcuHMI_AcuRev4100_VoltSwl_Hys_031(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_VoltSwl_Hys_031
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_event_waveform(app_page)
        set_pq_hysteresis(app_page, SWL_IDX, value)
        saved = save_and_check(app_page)
        _log.info("[TC] VoltSwellHysteresis=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Voltage Swell Hysteresis={value}: 保存失败"
            verify_modbus(REG_VOLT_SWL_HYS, value, label=f"VoltSwellHysteresis({value})")
        else:
            assert not saved, f"Voltage Swell Hysteresis={value}: 期望校验失败但保存成功"


