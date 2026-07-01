"""
Testcase_AcuHMI_AcuvimIIW_ClassS_WfSampleRate_040
Waveform Sample Rate options
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403


# ── 测试数据 ──────────────────────────────────────────────────────────
_CASES = [
    ('16 sample/cycle', 0),
    ('32 sample/cycle', 1),
    ('64 sample/cycle', 2),
    ('128 sample/cycle', 3),
    ('256 sample/cycle', 4),
    ('512 sample/cycle', 5),
]


def test_Testcase_AcuHMI_AcuvimIIW_ClassS_WfSampleRate_040(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_ClassS_WfSampleRate_040
    参数: option, encoded
    """
    for option, encoded in _CASES:
        _log.info("[TC] WfSampleRate=%s", option)
        nav_to_class_s_event(app_page)
        set_dropdown(app_page, "Waveform Sample Rate", option)
        saved = save_and_check(app_page)
        assert saved, f"WfSampleRate={option}: save failed"
        cs_verify(REG_CS_SAMPLE_RATE, encoded, label=f"WfSampleRate({option})")