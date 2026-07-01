"""
Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Hys_030
VoltIntr Hys
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (1, True),
    (5, True),
    (10, True),
    (0, False),
    (11, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Hys_030(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_VoltIntr_Hys_030
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] VoltIntr Hys=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_event_hys(app_page, 2, value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"VoltIntr Hysteresis={value}: save failed"
            cs_verify(cs_reg("VoltIntr", CS_OFFSET_HYS), value,
                      label=f"VoltIntr_Hys({value})")
        else:
            assert not saved, f"VoltIntr Hysteresis={value}: expected validation failure"