"""
Testcase_AcuHMI_AcuvimIIW_018_ClassS_RatedVoltage
Rated Voltage boundary (10-690 V)
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    (10, True),
    (200, True),
    (690, True),
    (9, False),
    (691, False),
]


def test_Testcase_AcuHMI_AcuvimIIW_018_ClassS_RatedVoltage(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_018_ClassS_RatedVoltage
    参数: value, ok
    """
    for value, ok in _CASES:
        _log.info("[TC] RatedVoltage=%s ok=%s", value, ok)
        nav_to_class_s_event(app_page)
        set_input(app_page, "Rated Voltage", value)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"RatedVoltage={value}: save failed"
            cs_verify(REG_CS_VOLT_RATED, value, label=f"RatedVoltage({value})")
        else:
            assert not saved, f"RatedVoltage={value}: expected validation failure"