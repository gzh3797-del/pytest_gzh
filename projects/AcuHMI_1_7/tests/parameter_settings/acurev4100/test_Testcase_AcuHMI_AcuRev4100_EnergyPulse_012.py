"""
Testcase_AcuHMI_AcuRev4100_EnergyPulse_012
原始编号: TC_EnergyPulse_001 ~ TC_EnergyPulse_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：hmi_value, reg_value, ok ─────────────────────────────────────────
_CASES = [
    ('0.100', 100, True),                                     # TC_EnergyPulse_001
    ('1000.000', 1000000, True),                              # TC_EnergyPulse_002
    ('100000.000', 100000000, True),                          # TC_EnergyPulse_003
    ('0.099', 99, False),                                     # TC_EnergyPulse_004
    ('100000.001', 100000001, False),                         # TC_EnergyPulse_005
]


def test_Testcase_AcuHMI_AcuRev4100_EnergyPulse_012(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_EnergyPulse_012
    参数组: hmi_value, reg_value, ok
    """
    for (hmi_value, reg_value, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "Energy Pulse Constant", hmi_value)
        saved = save_and_check(app_page)
        _log.info("[TC] Energy Pulse Constant=%s  ok=%s  result=%s", hmi_value, ok, saved)
        if ok:
            assert saved, f"Energy Pulse Constant={hmi_value}: 保存失败"
            verify_modbus_32(REG_ENERGY_PULSE, reg_value, label=f"EnergyPulseConst({hmi_value})")
        else:
            assert not saved, f"Energy Pulse Constant={hmi_value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_ENERGY_PULSE)


