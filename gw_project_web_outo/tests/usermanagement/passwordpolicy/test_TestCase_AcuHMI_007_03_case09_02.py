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


# 用例编号：TestCase_AcuHMI_007_03_case09_02
# 用例标题：设置失败登录等待为40000，保存配置成功
# 测试步骤：
#   1. Failed Login Wait=40000，保存
# 预期结果：
#   1. 保存成功
def test_TestCase_AcuHMI_007_03_case09_02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    try:
        _nav_to_submenu(page, "Password Policy")
        page.get_by_placeholder("Enter Maximum Failed Attempts").fill("5")
        page.get_by_placeholder("Enter Failed Login Attempt Window").fill("300")
        page.get_by_placeholder("Enter Failed Login Wait").fill("40000")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        has_error = page.locator(".el-form-item__error").count() > 0 or \
                    page.locator(".el-message--error").count() > 0
        assert not has_error, \
            "Failed Login Wait=40000 保存应成功，不应出现错误"
    finally:
        _restore_lockout_policy(page)
