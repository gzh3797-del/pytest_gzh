"""
Testcase_AcuHMI_AcuRev4100_Backlight_001
原始编号: TC_Backlight_001 ~ TC_Backlight_004
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (0, True),                                                # TC_Backlight_001
    (60, True),                                               # TC_Backlight_002
    (120, True),                                              # TC_Backlight_003
    (121, False),                                             # TC_Backlight_004
]


def test_Testcase_AcuHMI_AcuRev4100_Backlight_001(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_Backlight_001
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "Backlight", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Backlight=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Backlight={value}: 保存失败"
            verify_modbus(REG_BACKLIGHT, value, label="LCD Backlight Time")
        else:
            assert not saved, f"Backlight={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_BACKLIGHT)


