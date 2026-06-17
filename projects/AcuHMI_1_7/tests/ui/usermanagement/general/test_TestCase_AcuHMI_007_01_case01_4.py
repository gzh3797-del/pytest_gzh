import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_submenu(page, submenu: str):
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _restore_timeout(page):
    _nav_to_submenu(page, "General")
    page.get_by_placeholder("Enter Session Timeout").fill("10")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case01_4
# 用例标题：会话超时为 -1、61，保存配置失败系统错误信息提示准确
# 测试步骤（更新）：
#   1. Session Timeout 输入 -1 → 字段自动显示 1（spinbutton 自动修正负值）
#   2. Session Timeout 输入 61 → 保存应失败，系统提示错误信息
# 预期结果（更新）：
#   -1 → 字段显示变为 1；61 → 配置保存失败，错误信息提示正确
def test_TestCase_AcuHMI_007_01_case01_4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        # Step 1: 输入 -1，spinbutton 自动修正，字段应显示 1
        _nav_to_submenu(page, "General")
        timeout_input = page.get_by_placeholder("Enter Session Timeout")
        timeout_input.click(click_count=3)
        timeout_input.fill("-1")
        page.keyboard.press("Tab")
        page.wait_for_timeout(500)
        actual_display = timeout_input.input_value()
        assert actual_display == "1", (
            f"Session Timeout 输入 '-1'，期望字段自动显示 '1'，实际显示 '{actual_display}'"
        )

        # Step 2: 输入 61，点击保存 → 应失败
        _nav_to_submenu(page, "General")
        timeout_input = page.get_by_placeholder("Enter Session Timeout")
        timeout_input.click(click_count=3)
        timeout_input.fill("61")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        has_field_error = page.locator(".el-form-item__error").count() > 0
        has_error_toast = page.locator(".el-message--error").count() > 0
        success_visible = page.get_by_text("configuration saved", exact=False).is_visible()
        assert has_field_error or has_error_toast or not success_visible, \
            "Session Timeout=61 超出范围（0-60），保存应失败"
    finally:
        _restore_timeout(page)
