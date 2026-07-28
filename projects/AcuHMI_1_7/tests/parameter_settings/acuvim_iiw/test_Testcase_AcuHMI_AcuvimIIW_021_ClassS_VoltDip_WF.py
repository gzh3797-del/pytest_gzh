"""
Testcase_AcuHMI_AcuvimIIW_021_ClassS_VoltDip_WF
VoltDip WF
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (0, False),
    (1, True),
]


def test_Testcase_AcuHMI_AcuvimIIW_021_ClassS_VoltDip_WF(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_021_ClassS_VoltDip_WF
    参数: encoded, enable
    """
    for encoded, enable in _CASES:
        _log.info("[TC] VoltDip WF=%s", enable)
        nav_to_class_s_event(app_page)
        set_event_wf(app_page, 0, enable)
        saved = save_and_check(app_page)
        assert saved, f"VoltDip WF={enable}: save failed"
        cs_verify(cs_reg("VoltDip", CS_OFFSET_WF), encoded,
                  label=f"VoltDip_WF({enable})")