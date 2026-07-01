"""
Testcase_AcuHMI_AcuvimIIW_ClassS_VoltDip_Thr_026
VoltDip Thr
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (10, True),
    (50, True),
    (90, True),
    (9, False),
    (91, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_VoltDip_Thr_026(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_VoltDip_Thr_026
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] VoltDip Thr=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_event_thr(app_page, 0, value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"VoltDip Threshold={value}: save failed"
            cs_verify(cs_reg("VoltDip", CS_OFFSET_THR), value,
                      label=f"VoltDip_Thr({value})")
        else:
            assert not saved, f"VoltDip Threshold={value}: expected validation failure"