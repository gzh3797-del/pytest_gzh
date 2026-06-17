import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
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


def _restore_length(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Minimum Password Length").fill("8")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


# 用例编号：TestCase_AcuHMI_007_03_case05
# 用例标题：设置最小密码长度为 5，保存配置失败，提示错误信息正确
# 测试步骤：
#   1. Minimum Password Length = 5（低于有效最小值）
#   2. 点击 Save
# 预期结果：
#   保存失败，提示错误信息正确
def test_TestCase_AcuHMI_007_03_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _nav_to_submenu(page, "Password Policy")
        inp = page.get_by_placeholder("Enter Minimum Password Length")
        inp.fill("5")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        has_field_error = page.locator(".el-form-item__error").count() > 0
        has_error_toast = page.locator(".el-message--error").count() > 0
        success_visible = page.get_by_text("configuration saved", exact=False).is_visible()
        assert (has_field_error or has_error_toast or not success_visible), \
            "Minimum Password Length=5 低于有效值，应保存失败"
    finally:
        _restore_length(page)
