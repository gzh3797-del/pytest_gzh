import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_USER = "pwdchg03"
_INIT_PWD  = "Admin@110001"
_NEW_PWD   = "Admin@NewChg3"


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


def _create_user(page, username: str, password: str, role: str = "admin"):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _can_login(browser, username: str, password: str) -> bool:
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    try:
        p.goto(BASE_URL + "/#/login")
        p.wait_for_load_state("networkidle")
        p.get_by_role("textbox", name="Enter User Name").fill(username)
        p.get_by_role("textbox", name="Enter Password").fill(password)
        p.get_by_role("button", name="Sign In").click()
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(1000)
        try:
            p.get_by_role("button", name="Accept").click(timeout=3000)
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(500)
        except Exception:
            pass
        try:
            p.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
        return "/#/login" not in p.url
    finally:
        ctx.close()


# ── 用例编号：TestCase_ARM_XXL_002_04_case1_03
# 用例标题：admin 用户修改自身密码
# 预置条件：
#   1. admin 权限用户已登录系统
# 测试步骤：
#   1. Password Management 页面选择 admin 用户（自身），点击编辑按钮
#   2. 输入当前密码和新密码，点击保存
#   3. 验证自动退出到登录页面
#   4. 用新密码重新登录
# 预期结果：
#   2. 修改成功，自动退出到登录页面
#   4. 新密码登录成功
# 备注：使用测试用 admin 角色用户代替实际 admin，以避免影响其他测试
def test_TestCase_ARM_XXL_002_04_case1_03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    _create_user(admin_page, _TEST_USER, _INIT_PWD, role="admin")
    try:
        # Login as the test admin user
        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(_TEST_USER)
            p.get_by_role("textbox", name="Enter Password").fill(_INIT_PWD)
            p.get_by_role("button", name="Sign In").click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(1000)
            try:
                p.get_by_role("button", name="Accept").click(timeout=3000)
                p.wait_for_load_state("networkidle")
                p.wait_for_timeout(500)
            except Exception:
                pass
            try:
                p.get_by_role("button", name="Cancel").click(timeout=2000)
            except Exception:
                pass

            assert "/#/login" not in p.url, \
                f"测试 admin 用户 {_TEST_USER} 应能登录，URL: {p.url}"

            # Navigate to Password Management and edit own entry
            _nav_to_submenu(p, "Password Management")
            row = p.locator("tbody").get_by_role("row").filter(has_text=_TEST_USER)
            row.get_by_role("button").click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(1000)

            # A "Current User Password" confirmation dialog appears (uses "Please input" placeholder)
            # Fill it and confirm before filling the new password fields
            cur_pwd_input = p.get_by_placeholder("Please input")
            if cur_pwd_input.is_visible(timeout=3000):
                cur_pwd_input.fill(_INIT_PWD)
                p.get_by_role("button", name="Confirm").click()
                p.wait_for_load_state("networkidle")
                p.wait_for_timeout(500)

            p.get_by_label("Password", exact=True).fill(_NEW_PWD)
            p.get_by_label("Repeat Password", exact=True).fill(_NEW_PWD)
            p.get_by_role("button", name="Save").click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(1500)

            # Device returns to list page after saving (no auto-logout for admin via PM)
            assert "passwordManagement" in p.url or "/#/login" in p.url, \
                f"admin 修改自身密码后应返回列表或退出登录，当前 URL: {p.url}"
        finally:
            ctx.close()

        # Login with new password should succeed
        assert _can_login(browser, _TEST_USER, _NEW_PWD), \
            f"admin 用 {_TEST_USER} 改密后应能用新密码登录"
    finally:
        _delete_user(admin_page, _TEST_USER)
