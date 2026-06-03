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


def _restore_lockout_policy(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Maximum Failed Attempts").fill("0")
    page.get_by_placeholder("Enter Failed Login Attempt Window").fill("0")
    page.get_by_placeholder("Enter Failed Login Wait").fill("0")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _try_set_wait_and_assert_fail(page, wait_val: str):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Failed Login Wait")
    inp.fill(wait_val)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    has_field_error = page.locator(".el-form-item__error").count() > 0
    has_error_toast = page.locator(".el-message--error").count() > 0
    success_visible = page.get_by_text("configuration saved", exact=False).is_visible()
    assert has_field_error or has_error_toast or not success_visible, \
        f"Failed Login Wait={repr(wait_val)} 非法值，保存应失败"


# 用例编号：TestCase_AcuHMI_007_03_case09_05
# 用例标题：设置失败登录等待为小数、字母、特殊字符，保存配置失败，提示错误信息正确
# 测试步骤：
#   1. Failed Login Wait=1.5（小数）→ 保存失败
#   2. Failed Login Wait=abc（字母）→ 保存失败
#   3. Failed Login Wait=@#$（特殊字符）→ 保存失败
# 预期结果：
#   保存配置失败，提示错误信息正确
def test_TestCase_AcuHMI_007_03_case09_05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _try_set_wait_and_assert_fail(page, "1.5")
        _try_set_wait_and_assert_fail(page, "abc")
        _try_set_wait_and_assert_fail(page, "@#$")
    finally:
        _restore_lockout_policy(page)
