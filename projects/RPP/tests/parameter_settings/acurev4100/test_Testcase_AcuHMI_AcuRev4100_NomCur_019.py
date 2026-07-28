"""
Testcase_AcuHMI_AcuRev4100_NomCur_019
原始编号: TC_NomCur_001 ~ TC_NomCur_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (1, True),                                                # TC_NomCur_001
    (1000, True),                                             # TC_NomCur_002
    (50000, True),                                            # TC_NomCur_003
    (0, False),                                               # TC_NomCur_004
    (50001, False),                                           # TC_NomCur_005
]


def test_Testcase_AcuHMI_AcuRev4100_NomCur_019(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_NomCur_019
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "Nominal Current", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Nominal Current=%s  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Nominal Current={value}: 保存失败"
            verify_modbus(REG_NOMINAL_CURRENT, value, label="Nominal Current")
        else:
            assert not saved, f"Nominal Current={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_NOM_CUR)


