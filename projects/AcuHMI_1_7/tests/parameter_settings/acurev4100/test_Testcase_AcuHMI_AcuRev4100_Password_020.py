"""
Testcase_AcuHMI_AcuRev4100_Password_020
原始编号: TC_Password_001 ~ TC_Password_006
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_4100 import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：value, reg_val, ok ─────────────────────────────────────────
_CASES = [
    ('0000', 0, True),                                        # TC_Password_001
    ('1234', 1234, True),                                     # TC_Password_002
    ('9999', 9999, True),                                     # TC_Password_003
    ('10000', None, False),                                   # TC_Password_004
    ('123', None, False),                                     # TC_Password_005
    ('0', None, False),                                       # TC_Password_006
]


def test_Testcase_AcuHMI_AcuRev4100_Password_020(app_page):
    """
    Testcase_AcuHMI_AcuRev4100_Password_020
    参数组: value, reg_val, ok
    """
    for (value, reg_val, ok) in _CASES:
        nav_to_general(app_page)
        set_input(app_page, "Password", value)
        saved = save_and_check(app_page)
        _log.info("[TC] Password=%r  ok=%s  result=%s", value, ok, saved)
        if ok:
            assert saved, f"Password={value!r}: 保存失败"
            verify_modbus(REG_PASSWORD, reg_val, label=f"Password({value})")
            # 恢复默认值 "0000"，避免影响后续测试登录
            nav_to_general(app_page)
            set_input(app_page, "Password", "0000")
            save_and_check(app_page)
        else:
            assert not saved, f"Password={value!r}: 期望校验失败但保存成功"


