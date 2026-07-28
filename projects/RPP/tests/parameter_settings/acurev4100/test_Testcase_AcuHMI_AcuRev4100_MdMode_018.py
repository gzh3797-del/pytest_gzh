"""
Testcase_AcuHMI_AcuRev4100_MdMode_018
原始编号: TC_MdMode_001 ~ TC_MdMode_003
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, encoded ─────────────────────────────────────────
_CASES = [
    ('Manual Reset', 0),                                      # TC_MdMode_001
    ('Auto Reset', 1),                                        # TC_MdMode_002
    ('Disable', 2),                                           # TC_MdMode_003
]


def test_Testcase_AcuHMI_AcuRev4100_MdMode_018(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_MdMode_018
    参数组: option, encoded
    """
    for (option, encoded) in _CASES:
        nav_to_general(app_page)
        set_dropdown(app_page, "Mode", option)
        saved = save_and_check(app_page)
        assert saved, f"MaxDemand Clear Mode={option}: 保存失败"
        verify_modbus(REG_MD_CLEAR_MODE, encoded, label=f"MaxDemandClearMode({option})")


