"""
Testcase_AcuHMI_AcuRev4100_DmndInt_009
原始编号: TC_DmndInt_001 ~ TC_DmndInt_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_DmndInt_001
    (15, True),                                               # TC_DmndInt_002
    (30, True),                                               # TC_DmndInt_003
    (0, False),                                               # TC_DmndInt_004
    (31, False),                                              # TC_DmndInt_005
]


def test_Testcase_AcuHMI_AcuRev4100_DmndInt_009(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_DmndInt_009
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "Demand Interval", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Demand Interval=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Demand Interval={value}: 保存失败"
            verify_modbus(REG_DEMAND_INTERVAL, value, label="Demand Interval")
        else:
            assert not saved, f"Demand Interval={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_DMND_INT)


