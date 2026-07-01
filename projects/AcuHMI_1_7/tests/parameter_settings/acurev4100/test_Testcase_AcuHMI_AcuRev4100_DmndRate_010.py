"""
Testcase_AcuHMI_AcuRev4100_DmndRate_010
原始编号: TC_DmndRate_001 ~ TC_DmndRate_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_DmndRate_001
    (15, True),                                               # TC_DmndRate_002
    (30, True),                                               # TC_DmndRate_003
    (0, False),                                               # TC_DmndRate_004
    (31, False),                                              # TC_DmndRate_005
]


def test_Testcase_AcuHMI_AcuRev4100_DmndRate_010(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_DmndRate_010
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        step("前置：Demand Method = Sliding（避免 Update Rate 置灰）")
        nav_to_general(app_page)
        set_dropdown(app_page, "Demand Method", "Sliding")
        app_page.wait_for_timeout(500)
        set_input(app_page, "Demand Update Rate", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Demand Update Rate=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Demand Update Rate={value}: 保存失败"
            verify_modbus(REG_DEMAND_SUBINT, value, label="Demand Update Rate")
        else:
            assert not saved, f"Demand Update Rate={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_DMND_INT)


