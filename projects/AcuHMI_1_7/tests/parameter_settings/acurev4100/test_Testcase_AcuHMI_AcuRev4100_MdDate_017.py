"""
Testcase_AcuHMI_AcuRev4100_MdDate_017
原始编号: TC_MdDate_001 ~ TC_MdDate_002
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：day_option, expected_day ─────────────────────────────────────────
_CASES = [
    ('01', 1),                                                # TC_MdDate_001
    ('31', 31),                                               # TC_MdDate_002
]


def test_Testcase_AcuHMI_AcuRev4100_MdDate_017(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_MdDate_017
    参数组: day_option, expected_day
    """
    for (day_option, expected_day) in _CASES:
        nav_to_general(app_page)
        step("前置：MaxDemand Clear Mode = Auto Reset")
        set_dropdown(app_page, "Mode", "Auto Reset")
        app_page.wait_for_timeout(300)
        set_dropdown(app_page, "Auto Reset Date", day_option)
        saved = save_and_check(app_page)
        assert saved, f"Auto Reset Date={day_option}: 保存失败"
        regs = modbus_read(REG_MD_CLEAR_DATE, count=1)
        actual_day = (regs[0] >> 8) & 0xFF
        _log.info("[VERIFY] AutoResetDate  expected_day=%-4s  actual_day=%-4s  reg=0x%04X  → %s",
                  expected_day, actual_day, regs[0], "✓ PASS" if actual_day == expected_day else "✗ FAIL")
        assert actual_day == expected_day, (
            f"Auto Reset Date: expected day={expected_day}, got {actual_day} (reg={regs[0]})"
        )


