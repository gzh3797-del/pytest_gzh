"""
Testcase_AcuHMI_AcuvimIIW_EnergyType_047
原始编号: TC_EnergyType_001 ~ TC_EnergyType_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Fundamental', 0),                                       # TC_EnergyType_001
    ('Fundament+Harmonics', 1),                               # TC_EnergyType_002
]


def test_Testcase_AcuHMI_AcuvimIIW_EnergyType_047(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_EnergyType_047
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════EnergyType=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Advanced")
        set_dropdown(app_page, "Energy Type", option)
        saved = save_and_check(app_page)
        assert saved, f"Energy Type={option}: 保存失败"
        _log.info("[EXPECT  ] EnergyType reg = %s", encoded)
        verify_reg(REG_ENERGY_CALC, encoded, label=f"EnergyType({option})")


