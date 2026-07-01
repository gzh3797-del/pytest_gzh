# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_06_case01
类别：UI — 页面布局与标签验证
功能：Data Log Parameter Config 页面标题拼写、面板标签、按钮文字均符合规格
"""
from datalog_page import DataLogParamConfigPage

CASE_ID = "TestCase_AcuHMI_003_06_case01"

_EXPECTED_PAGE_TITLE  = "Data Log Parameter Config"
_EXPECTED_PANEL_LABELS = ["Not Selected", "Selected"]
_EXPECTED_BUTTONS      = ["All", "Clear", "Save"]


def _find_exact_text(page, text: str) -> list:
    xpath = (
        f"//*["
        f"  normalize-space(.)='{text}'"
        f"  and not(self::script)"
        f"  and not(self::style)"
        f"  and not(self::input)"
        f"  and not(self::textarea)"
        f"]"
    )
    return page.locator(f"xpath={xpath}").all()


def test_case(driver):
    page = DataLogParamConfigPage(driver)
    page.navigate()

    # 1. 页面标题文本验证
    title_els = _find_exact_text(driver, _EXPECTED_PAGE_TITLE)
    assert title_els, (
        f"[{CASE_ID}] 未找到文本 '{_EXPECTED_PAGE_TITLE}'。"
        f"请确认面包屑/菜单中单词间有空格，拼写无误"
    )

    # 2. 双面板区域标题拼写验证
    missing_panels = []
    for panel_text in _EXPECTED_PANEL_LABELS:
        els = _find_exact_text(driver, panel_text)
        if not els:
            missing_panels.append(panel_text)
    assert not missing_panels, (
        f"[{CASE_ID}] 面板标题缺失或拼写错误：{missing_panels}"
    )

    # 3. 操作按钮文字验证
    missing_buttons = []
    for btn_text in _EXPECTED_BUTTONS:
        els = driver.locator(
            f"xpath=//button[.//span[normalize-space(.)='{btn_text}']]"
            f" | //button[normalize-space(.)='{btn_text}']"
        ).all()
        if not els:
            missing_buttons.append(btn_text)
    assert not missing_buttons, (
        f"[{CASE_ID}] 按钮缺失或文字错误：{missing_buttons}"
    )
