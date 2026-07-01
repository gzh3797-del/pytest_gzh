"""
Testcase_AcuHMI_AcuRev4100_LedPulse_014
原始编号: TC_LedPulse_001 ~ TC_LedPulse_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：subcategory, param_name, expected_reg ─────────────────────────────────────────
_CASES = [
    ('Switch', 'Disable', 0),                                 # TC_LedPulse_001
    ('Input Channel Energy', 'Input Channel 1 active energy import', 37),  # TC_LedPulse_002
    ('Input Channel Energy', 'Input Channel 1 reactive energy net', 43),  # TC_LedPulse_003
    ('System Energy', 'Phase A active energy import', 1),     # TC_LedPulse_004
    ('User Channel Energy', 'User Channel 1 active energy import', 253),  # TC_LedPulse_005
]


def test_Testcase_AcuHMI_AcuRev4100_LedPulse_014(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_LedPulse_014
    参数组: subcategory, param_name, expected_reg
    """
    for (subcategory, param_name, expected_reg) in _CASES:
        nav_to_general(app_page)
        open_parameter_selector(app_page)
        select_led_param(app_page, subcategory, param_name)
        saved = save_and_check(app_page)
        assert saved, f"LED Pulse Parameter={param_name!r}: 保存失败"
        verify_modbus(REG_LED_PULSE_PARAM, expected_reg, label=f"LED Pulse Param ({param_name})")


