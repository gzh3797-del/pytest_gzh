"""
Testcase_AcuHMI_AcuvimIIW_006_CT42
原始编号: TC_CT42_001 ~ TC_CT42_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('5A', 5),                                                # TC_CT42_001
    ('1A', 1),                                                # TC_CT42_002
]


def test_Testcase_AcuHMI_AcuvimIIW_006_CT42(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_006_CT42
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════CT42=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_dropdown(app_page, "CT42", option)
        saved = save_and_check(app_page)
        assert saved, f"CT42={option}: 保存失败"
        _log.info("[EXPECT  ] CT42 reg = %s", encoded)
        verify_reg(REG_CT42, encoded, label=f"CT42({option})")


