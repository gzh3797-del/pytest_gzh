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


def _restore_grace(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Grace Period").fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


# 用例编号：TestCase_AcuHMI_007_03_case06_2
# 用例标题：设置宽限期为 30000，保存配置成功
# 测试步骤：
#   1. Grace Period = 30000
#   2. 点击 Save
# 预期结果：
#   保存成功（30000 为有效值）
def test_TestCase_AcuHMI_007_03_case06_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _nav_to_submenu(page, "Password Policy")
        inp = page.get_by_placeholder("Enter Grace Period")
        inp.fill("30000")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        has_field_error = page.locator(".el-form-item__error").count() > 0
        has_error_toast = page.locator(".el-message--error").count() > 0
        assert not has_field_error and not has_error_toast, \
            "Grace Period=30000 是有效值，应保存成功"
    finally:
        _restore_grace(page)
