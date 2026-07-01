"""
Testcase_AcuHMI_AcuvimIIW_I4_048
原始编号: TC_I4_001 ~ TC_I4_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('calculated', 0),                                        # TC_I4_001
    ('measured', 1),                                          # TC_I4_002
]


def test_Testcase_AcuHMI_AcuvimIIW_I4_048(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_I4_048
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════I4Method=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_dropdown(app_page, "I4 Method", option)
        saved = save_and_check(app_page)
        assert saved, f"I4 Method={option}: 保存失败"
        _log.info("[EXPECT  ] I4Method reg = %s", encoded)
        verify_reg(REG_I4_METHOD, encoded, label=f"I4Method({option})")


