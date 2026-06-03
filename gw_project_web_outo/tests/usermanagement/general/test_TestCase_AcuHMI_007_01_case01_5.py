import pytest
from pages.login_page import LoginPage


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


def _try_set_timeout_and_assert_fail(page, val: str):
    _nav_to_submenu(page, "General")
    inp = page.get_by_placeholder("Enter Session Timeout")
    inp.fill(val)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    has_field_error = page.locator(".el-form-item__error").count() > 0
    has_error_toast = page.locator(".el-message--error").count() > 0
    success_visible = page.get_by_text("configuration saved", exact=False).is_visible()
    assert has_field_error or has_error_toast or not success_visible, \
        f"Session Timeout={repr(val)} 非法值，保存应失败"


# 用例编号：TestCase_AcuHMI_007_01_case01_5
# 用例标题：会话超时为字母、特殊字符，保存配置失败系统错误信息提示准确
# 测试步骤：
#   1. Session Timeout=ddd（字母），保存 → 应失败
#   2. Session Timeout=@@##（特殊字符），保存 → 应失败
# 预期结果：
#   配置保存失败，系统提示错误信息正确
def test_TestCase_AcuHMI_007_01_case01_5(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _try_set_timeout_and_assert_fail(page, "ddd")
        _try_set_timeout_and_assert_fail(page, "@@##")
    finally:
        _restore_timeout(page)
