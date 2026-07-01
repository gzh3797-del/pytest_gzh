"""
Testcase_AcuHMI_AcuvimIIW_EnergyReading_046
原始编号: TC_EnergyReading_001 ~ TC_EnergyReading_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Primary', 0),                                           # TC_EnergyReading_001
    ('Secondary', 1),                                         # TC_EnergyReading_002
    ('Primary in 0.01kWh', 2),                                # TC_EnergyReading_003
]


def test_Testcase_AcuHMI_AcuvimIIW_EnergyReading_046(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_EnergyReading_046
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════EnergyReading=%s  expect_encoded=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Advanced")
        set_dropdown(app_page, "Energy Reading", option)
        saved = save_and_check(app_page)
        assert saved, f"Energy Reading={option}: 保存失败"


