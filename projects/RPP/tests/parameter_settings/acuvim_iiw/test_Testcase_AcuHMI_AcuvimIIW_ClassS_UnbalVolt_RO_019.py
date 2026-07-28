"""
Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_RO_019
UnbalVolt RO
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    ('No Output', 0),
    ('1-1 RO1', 1),
    ('3-2 RO1', 64),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_RO_019(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_RO_019
    参数: option, encoded
    """
    for option, encoded in _CASES:
        _log.info("[TC] UnbalVolt RO=%s", option)
        nav_to_class_s_event(app_page)
        set_event_ro(app_page, 3, option)
        saved = save_and_check(app_page)
        assert saved, f"UnbalVolt RO={option}: save failed"
        cs_verify(cs_reg("UnbalVolt", CS_OFFSET_RO), encoded,
                  label=f"UnbalVolt_RO({option})")