"""
Testcase_AcuHMI_AcuvimIIW_026_ClassS_UnbalCurr_Enable
UnbalCurr Enable
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (0, False),
    (1, True),
]


def test_Testcase_AcuHMI_AcuvimIIW_026_ClassS_UnbalCurr_Enable(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_026_ClassS_UnbalCurr_Enable
    参数: encoded, enable
    """
    for encoded, enable in _CASES:
        _log.info("[TC] UnbalCurr Enable=%s", enable)
        nav_to_class_s_event(app_page)
        set_event_enable(app_page, 5, enable)
        saved = save_and_check(app_page)
        assert saved, f"UnbalCurr Enable={enable}: save failed"
        cs_verify(cs_reg("UnbalCurr", CS_OFFSET_ENABLE), encoded,
                  label=f"UnbalCurr_Enable({enable})")