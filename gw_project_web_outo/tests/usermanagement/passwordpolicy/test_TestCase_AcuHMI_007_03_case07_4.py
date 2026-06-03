import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
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


def _restore_failed_login_settings(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Maximum Failed Attempts").fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _set_and_verify_fail(page, value: str):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Maximum Failed Attempts")
    inp.fill(value)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    has_field_error = page.locator(".el-form-item__error").count() > 0
    has_error_toast = page.locator(".el-message--error").count() > 0
    success_visible = page.get_by_text("configuration saved", exact=False).is_visible()
    assert (has_field_error or has_error_toast or not success_visible), \
        f"Maximum Failed Attempts={value!r} 应保存失败"


# 用例编号：TestCase_AcuHMI_007_03_case07_4
# 用例标题：设置最大失败尝试次数为 -1、31，保存配置失败，提示错误信息正确
# 测试步骤：
#   依次将 Maximum Failed Attempts 设置为 "-1"、"31"，点击 Save
# 预期结果：
#   各次保存均失败，提示错误信息正确
def test_TestCase_AcuHMI_007_03_case07_4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _set_and_verify_fail(page, "-1")
        _set_and_verify_fail(page, "31")
    finally:
        _restore_failed_login_settings(page)
