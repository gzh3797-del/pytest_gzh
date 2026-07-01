"""
Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_WF_021
UnbalVolt WF
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (0, False),
    (1, True),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_WF_021(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_WF_021
    参数: encoded, enable
    """
    for encoded, enable in _CASES:
        _log.info("[TC] UnbalVolt WF=%s", enable)
        nav_to_class_s_event(app_page)
        set_event_wf(app_page, 3, enable)
        saved = save_and_check(app_page)
        assert saved, f"UnbalVolt WF={enable}: save failed"
        cs_verify(cs_reg("UnbalVolt", CS_OFFSET_WF), encoded,
                  label=f"UnbalVolt_WF({enable})")