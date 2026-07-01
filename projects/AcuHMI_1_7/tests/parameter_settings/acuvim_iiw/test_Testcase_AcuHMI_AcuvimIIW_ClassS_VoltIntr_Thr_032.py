"""
Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Thr_032
VoltIntr Thr
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (5, True),
    (7, True),
    (10, True),
    (4, False),
    (11, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Thr_032(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Thr_032
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] VoltIntr Thr=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_event_thr(app_page, 2, value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"VoltIntr Threshold={value}: save failed"
            cs_verify(cs_reg("VoltIntr", CS_OFFSET_THR), value,
                      label=f"VoltIntr_Thr({value})")
        else:
            assert not saved, f"VoltIntr Threshold={value}: expected validation failure"