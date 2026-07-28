"""
Testcase_AcuHMI_AcuRev4100_LedWidth_015
原始编号: TC_LedWidth_001 ~ TC_LedWidth_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (20, True),                                               # TC_LedWidth_001
    (60, True),                                               # TC_LedWidth_002
    (100, True),                                              # TC_LedWidth_003
    (19, False),                                              # TC_LedWidth_004
    (101, False),                                             # TC_LedWidth_005
]


def test_Testcase_AcuHMI_AcuRev4100_LedWidth_015(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_LedWidth_015
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "LED Pulse Width", value)
        saved = save_and_check(app_page)
        _log.info("[TC] LED Pulse Width=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"LED Pulse Width={value}: 保存失败"
            verify_modbus(REG_LED_PULSE_WIDTH, value, label="LED Pulse Width")
        else:
            assert not saved, f"LED Pulse Width={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_LED_WIDTH)


