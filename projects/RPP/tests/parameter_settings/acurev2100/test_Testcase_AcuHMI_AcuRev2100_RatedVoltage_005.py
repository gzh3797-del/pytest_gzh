"""
Testcase_AcuHMI_AcuRev2100_RatedVoltage_005
原始编号: TC_Basic_001 ~ TC_Basic_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok, err_msg ─────────────────────────────────────────
_CASES = [
    (10, True, None),                                         # TC_Basic_001
    (200, True, None),                                        # TC_Basic_002
    (400, True, None),                                        # TC_Basic_003
    (9, False, _ERR_RATED_VOLTAGE),                           # TC_Basic_004
    (401, False, _ERR_RATED_VOLTAGE),                         # TC_Basic_005
]


def test_Testcase_AcuHMI_AcuRev2100_RatedVoltage_005(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_RatedVoltage_005
    参数组: value, ok, err_msg
    """
    for (value, ok, err_msg) in _CASES:
        nav_to_general(app_page, "Basic")
        set_input(app_page, "Rated Voltage", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Rated Voltage=%s  ok_expected=%s  save_result=%s", value, ok, saved)
        if ok:
            assert saved, f"Rated Voltage={value}: 期望成功但出现错误"
            verify_modbus(REG_RATED_VOLTAGE, value, label="Rated Voltage")
        else:
            assert not saved, f"Rated Voltage={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, err_msg)


