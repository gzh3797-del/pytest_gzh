"""
Testcase_AcuHMI_AcuRev4100_RPM_023
原始编号: TC_RPM_001 ~ TC_RPM_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Generic', 0),                                           # TC_RPM_001
    ('True', 1),                                              # TC_RPM_002
]


def test_Testcase_AcuHMI_AcuRev4100_RPM_023(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_RPM_023
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Reactive Power Calculation Method", option)
        saved = save_and_check(app_page)
        assert saved, f"Reactive Power Method={option}: 保存失败"
        verify_modbus(REG_REACTIVE_METHOD, encoded, label=f"ReactiveMethod({option})")


