"""
Testcase_AcuHMI_AcuvimIIW_PT_050
原始编号: TC_PT_001 ~ TC_PT_011
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_iiw import *  # noqa: F401, F403  # noqa: F401, F403


# ── 测试数据：pt1, pt2, ok, note, err_kw ─────────────────────────────────────────
_CASES = [
    (51, 50, True, 'PT1 最小有效值', None),                        # TC_PT_001
    (480, 100, True, 'PT1 典型值', None),                        # TC_PT_002
    (1000000, 50, True, 'PT1 最大值', None),                     # TC_PT_003
    (49, 50, False, 'PT1 低于下限', _ERR_PT1),                    # TC_PT_004
    (1000001, 50, False, 'PT1 超出上限', _ERR_PT1),               # TC_PT_005
    (1000, 50, True, 'PT2 最小值', None),                        # TC_PT_006
    (1000, 400, True, 'PT2 最大值', None),                       # TC_PT_007
    (1000, 49, False, 'PT2 低于下限', _ERR_PT2),                  # TC_PT_008
    (1000, 401, False, 'PT2 超出上限', _ERR_PT2),                 # TC_PT_009
    (500, 400, True, 'PT2 最大值 PT1>PT2', None),                # TC_PT_010
    (500, 500, False, 'PT2=500 超出上限(max=400)', _ERR_PT2),     # TC_PT_011
]


def test_Testcase_AcuHMI_AcuvimIIW_PT_050(app_page):
    """
    Testcase_AcuHMI_AcuvimIIW_PT_050
    参数组: pt1, pt2, ok, note, err_kw
    """
    for (pt1, pt2, ok, note, err_kw) in _CASES:
        _log.info("[TC] ════PT1=%-8s PT2=%-5s ok=%-5s  %s", pt1, pt2, ok, note)
        nav_to_general(app_page)
        click_tab(app_page, "Basic")
        set_input(app_page, "PT1", pt1)
        set_input(app_page, "PT2", pt2)
        saved = save_and_check(app_page)
        if ok:
            assert saved, f"PT1={pt1}, PT2={pt2} ({note}): 保存失败"
            # PT1/PT2 寄存器均以 0.1 为单位存储（实测 display_value × 10）
            _log.info("[EXPECT  ] PT1 reg = %s × 10 = %s", pt1, pt1 * 10)
            verify_reg_32(REG_PT1_HIGH, pt1 * 10, label="PT1")
            _log.info("[EXPECT  ] PT2 reg = %s × 10 = %s", pt2, pt2 * 10)
            verify_reg(REG_PT2, pt2 * 10, label="PT2")
        else:
            assert not saved, f"PT1={pt1}, PT2={pt2} ({note}): 期望校验失败但保存成功"
            assert_field_error(app_page, err_kw)


