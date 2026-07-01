"""
Testcase_AcuHMI_AcuRev2100_CtChannels_003
原始编号: TC_Basic_014 ~ TC_Basic_018
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_2100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, ok ─────────────────────────────────────────
_CASES = [
    (5, True),                                                # TC_Basic_014
    (25000, True),                                            # TC_Basic_015
    (50000, True),                                            # TC_Basic_016
    (4, False),                                               # TC_Basic_017
    (50001, False),                                           # TC_Basic_018
]


def test_Testcase_AcuHMI_AcuRev2100_CtChannels_003(app_page):
    """
    Testcase_AcuHMI_AcuRev2100_CtChannels_003
    参数组: value, ok
    """
    for (value, ok) in _CASES:
        nav_to_general(app_page, "Basic")
        for lbl in CT_LABELS:
            set_input(app_page, lbl, value)
        saved = save_and_check(app_page)
        _log.info("[TC] CT channels=%s  ok_expected=%s  save_result=%s", value, ok, saved)
        if ok:
            assert saved, f"CT channels={value}: 保存失败"
            verify_modbus(REG_CT_FS_BASE, value, label="CT Ch1~18", count=18)
        else:
            assert not saved, f"CT channels={value}: 期望校验失败但保存成功"
            assert_field_error(app_page, _ERR_CT_CH)


