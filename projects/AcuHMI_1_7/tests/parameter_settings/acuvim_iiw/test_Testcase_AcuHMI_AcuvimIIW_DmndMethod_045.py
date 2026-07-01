"""
Testcase_AcuHMI_AcuvimIIW_DmndMethod_045
原始编号: TC_DmndMethod_001 ~ TC_DmndMethod_004
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Fixed block', 0),                                       # TC_DmndMethod_001
    ('Sliding block', 1),                                     # TC_DmndMethod_002
    ('thermal', 2),                                           # TC_DmndMethod_003
    ('Rolling block', 3),                                     # TC_DmndMethod_004
]


def test_Testcase_AcuHMI_AcuvimIIW_DmndMethod_045(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_DmndMethod_045
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        _log.info("[TC] ════DemandMethod=%s  expect_reg=%s", option, encoded)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_dropdown(app_page, "Sliding Window", option)
        saved = save_and_check(app_page)
        assert saved, f"Sliding Window={option}: 保存失败"
        _log.info("[EXPECT  ] DemandMethod reg = %s", encoded)
        verify_reg(REG_DEMAND_METHOD, encoded, label=f"DemandMethod({option})")


