import re
import pytest

from playwright.sync_api import expect

from config.settings import DEFAULT_USERNAME, DEFAULT_PASSWORD
from pages.login_page import LoginPage


# @pytest.mark.login
# @pytest.mark.smoke
# class TestLogin:
#
#     # def test_login_success(self, login_page: LoginPage):
#     #     login_page.open()
#     #     login_page.login(DEFAULT_USERNAME, DEFAULT_PASSWORD)
#     #     assert login_page.is_logged_in(), "Login failed -- logout button not visible after login"
#     #
#     # def test_login_wrong_password(self, login_page: LoginPage):
#     #     login_page.open()
#     #     login_page.login(DEFAULT_USERNAME, "wrong_password")
#     #     assert not login_page.is_logged_in(), "Should not be logged in with wrong password"
#     #
#     # def test_login_empty_credentials(self, login_page: LoginPage):
#     #     login_page.open()
#     #     login_page.login("", "")
#     #     assert not login_page.is_logged_in(), "Should not be logged in with empty credentials"


def test_useradd(login_page: LoginPage):
    # 登录
    login_page.open()
    login_page.login()

    # 导航到用户管理页面
    page = login_page.page
    page.locator("header span").filter(has_text="AcuHMI").click()
    page.get_by_text("User Management").click()
    page.get_by_role("menuitem", name="User Configuration").click()

    # 新增用户
    page.get_by_role("button", name="Add User").click()
    page.get_by_role("textbox", name="Username*").fill("test_user")
    page.get_by_role("textbox", name="Password*", exact=True).fill("Admin@12345678")
    page.get_by_role("textbox", name="Repeat Password*").fill("Admin@12345678")
    page.locator(".el-icon.el-select__caret > svg").click()
    page.get_by_role("option", name="view").click()
    page.get_by_role("button", name="Save").click()
    # 等待 Save 后弹窗完全消失，避免遮挡后续操作
    page.wait_for_selector(".el-overlay-message-box", state="hidden", timeout=10000)

    # 断言用户已添加
    expect(page.locator("tbody")).to_contain_text("test_user")
    expect(page.locator("tbody")).to_contain_text("view")

    # 删除用户
    page.get_by_role("button").filter(has_text=re.compile(r"^$")).nth(4).click()
    page.get_by_role("button", name="Yes, continue").click()

    # 断言用户已删除
    expect(page.locator("tbody")).not_to_contain_text("test_user")

def test_userlogin(login_page: LoginPage):
    login_page.open()
    login_page.login(DEFAULT_USERNAME, DEFAULT_PASSWORD)
    assert login_page.is_logged_in(), "登录失败：AcuHMI 菜单未出现"
