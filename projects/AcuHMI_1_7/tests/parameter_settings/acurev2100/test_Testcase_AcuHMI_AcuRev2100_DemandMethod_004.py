"""
Testcase_AcuHMI_AcuRev2100_DemandMethod_004
原始编号: TC_Basic_019 ~ TC_Basic_020
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded, sub_disabled ─────────────────────────────────────────
_CASES = [
    ('Rolling Window Demand', 1, False),                      # TC_Basic_019
    ('Fixed Window Demand', 2, True),                         # TC_Basic_020
]


def test_Testcase_AcuHMI_AcuRev2100_DemandMethod_004(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_DemandMethod_004
    参数组: option, encoded, sub_disabled
    """
    for (option, encoded, sub_disabled) in _CASES:
        nav_to_general(app_page, "Basic")
        set_dropdown(app_page, "Sliding Window", option)
        app_page.wait_for_timeout(600)
        actual_disabled = sub_interval_is_disabled(app_page)
        assert actual_disabled == sub_disabled, (
            f"Demand={option}: Sub-Interval 置灰期望={sub_disabled}, 实际={actual_disabled}"
        )
        saved = save_and_check(app_page)
        assert saved, f"Demand Method={option}: 保存失败"
        verify_modbus(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


