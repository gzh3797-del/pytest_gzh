"""
Testcase_AcuHMI_AcuvimIIW_ClassS_VoltSwl_Enable_035
VoltSwl Enable
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (0, False),
    (1, True),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_VoltSwl_Enable_035(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_VoltSwl_Enable_035
    参数: encoded, enable
    """
    for encoded, enable in _CASES:
        _log.info("[TC] VoltSwl Enable=%s", enable)
        nav_to_class_s_event(app_page)
        set_event_enable(app_page, 1, enable)
        saved = save_and_check(app_page)
        assert saved, f"VoltSwl Enable={enable}: save failed"
        cs_verify(cs_reg("VoltSwl", CS_OFFSET_ENABLE), encoded,
                  label=f"VoltSwl_Enable({enable})")