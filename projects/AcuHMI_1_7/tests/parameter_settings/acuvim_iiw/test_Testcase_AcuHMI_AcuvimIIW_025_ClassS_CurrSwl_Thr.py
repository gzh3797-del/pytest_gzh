"""
Testcase_AcuHMI_AcuvimIIW_025_ClassS_CurrSwl_Thr
CurrSwl Thr
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (110, True),
    (130, True),
    (150, True),
    (109, False),
    (151, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_025_ClassS_CurrSwl_Thr(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_025_ClassS_CurrSwl_Thr
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] CurrSwl Thr=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_event_thr(app_page, 4, value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"CurrSwl Threshold={value}: save failed"
            cs_verify(cs_reg("CurrSwl", CS_OFFSET_THR), value,
                      label=f"CurrSwl_Thr({value})")
        else:
            assert not saved, f"CurrSwl Threshold={value}: expected validation failure"