"""
Testcase_AcuHMI_AcuvimIIW_019_ClassS_RatedCurrent
Rated Current boundary (0.1-5 A)
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    ('0.1', True),
    ('3', True),
    ('5', True),
    ('0.09', False),
    ('5.1', False),
]


def test_Testcase_AcuHMI_AcuvimIIW_019_ClassS_RatedCurrent(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_019_ClassS_RatedCurrent
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] RatedCurrent=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_input(app_page, "Rated Current", value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"RatedCurrent={value}: save failed"
        else:
            assert not saved, f"RatedCurrent={value}: expected validation failure"