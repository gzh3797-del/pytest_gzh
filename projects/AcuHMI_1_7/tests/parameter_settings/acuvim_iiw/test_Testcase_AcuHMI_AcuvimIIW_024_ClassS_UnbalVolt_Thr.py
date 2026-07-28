"""
Testcase_AcuHMI_AcuvimIIW_024_ClassS_UnbalVolt_Thr
UnbalVolt Thr
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (5, True),
    (25, True),
    (50, True),
    (4, False),
    (51, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_024_ClassS_UnbalVolt_Thr(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_024_ClassS_UnbalVolt_Thr
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] UnbalVolt Thr=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_event_thr(app_page, 3, value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"UnbalVolt Threshold={value}: save failed"
            cs_verify(cs_reg("UnbalVolt", CS_OFFSET_THR), value,
                      label=f"UnbalVolt_Thr({value})")
        else:
            assert not saved, f"UnbalVolt Threshold={value}: expected validation failure"