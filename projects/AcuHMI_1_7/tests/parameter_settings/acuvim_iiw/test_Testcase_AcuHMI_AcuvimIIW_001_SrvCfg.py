"""
Testcase_AcuHMI_AcuvimIIW_001_SrvCfg
原始编号: TC_SrvCfg_001 ~ TC_SrvCfg_005
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：option, v_reg, c_reg ─────────────────────────────────────────
_CASES = [
    ('3 Element 4 Wire Wye (3LN-3CT)', 0, 0),                 # TC_SrvCfg_001
    ('3 Element 3 Wire Delta (3LL-3CT)', 3, 0),               # TC_SrvCfg_002
    ('2 Element 3 Wire Network (2LN-2CT)', 6, 2),             # TC_SrvCfg_003
    ('2 Element 3 Wire 1 Phase (1LL-2CT)', 4, 2),             # TC_SrvCfg_004
    ('1 Element 2 Wire (1LN-1CT)', 1, 1),                     # TC_SrvCfg_005
]


def test_Testcase_AcuHMI_AcuvimIIW_001_SrvCfg(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_001_SrvCfg
    参数组: option, v_reg, c_reg
    """
    for (option, v_reg, c_reg) in _CASES:
        _log.info("[TC] ════ServiceConfig=%r  expect VoltWiring=%s CurrWiring=%s", option, v_reg, c_reg)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_service_config(app_page, option)
        saved = save_and_check(app_page)
        assert saved, f"Service Config={option!r}: 保存失败"
        _log.info("[EXPECT  ] VoltWiring reg = %s", v_reg)
        verify_reg(REG_VOLT_WIRING, v_reg, label=f"VoltWiring({option})")
        _log.info("[EXPECT  ] CurrWiring reg = %s", c_reg)
        verify_reg(REG_CURR_WIRING, c_reg, label=f"CurrWiring({option})")


