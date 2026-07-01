"""
Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_DO_016
UnbalVolt DO
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    ('No Output', 0),
    ('2-1 DO1', 1),
    ('2-2 DO2', 4),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_DO_016(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_UnbalVolt_DO_016
    参数: option, encoded
    """
    for option, encoded in _CASES:
        _log.info("[TC] UnbalVolt DO=%s", option)
        nav_to_class_s_event(app_page)
        set_event_do(app_page, 3, option)
        saved = save_and_check(app_page)
        assert saved, f"UnbalVolt DO={option}: save failed"
        cs_verify(cs_reg("UnbalVolt", CS_OFFSET_DO), encoded,
                  label=f"UnbalVolt_DO({option})")