import pytest
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


def _create_user(page, username: str, password: str, role: str = "view"):
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


def _lock_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").first.click()  # Lock is the first button for non-admin rows
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_verify(browser, username: str, password: str) -> bool:
    """在独立浏览器上下文中尝试登录，返回是否登录成功。"""
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


# 用例编号：TestCase_AcuHMI_007_01_case04_2
# 用例标题：锁全部用户，登录所有用户
# 测试步骤（更新）：
#   1. 创建 uc042_1、uc042_2（view 角色）
#   2. 锁定全部非 admin 用户（admin 用户无法锁定）
#   3. 被锁定的用户尝试登录 → 应失败
#   4. admin 用户尝试登录 → 应成功
# 预期结果（更新）：
#   1. 除 admin 外，其他用户均无法登录成功
#   2. admin 用户可以正常登录
def test_TestCase_AcuHMI_007_01_case04_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    pwd = "Abc@12345"
    test_users = ["uc042_1", "uc042_2"]

    try:
        for username in test_users:
            _create_user(page, username, pwd, role="view")

        # Lock all non-admin users
        for username in test_users:
            _lock_user(page, username)

        # Step 1: Locked users must NOT be able to login
        for username in test_users:
            logged_in = _login_verify(browser, username, pwd)
            assert not logged_in, f"已锁定用户 {username} 不应能登录成功"

        # Step 2: Admin must still be able to login (admin cannot be locked)
        from projects.AcuHMI_1_7.settings import DEFAULT_USERNAME, DEFAULT_PASSWORD
        admin_logged_in = _login_verify(browser, DEFAULT_USERNAME, DEFAULT_PASSWORD)
        assert admin_logged_in, "admin 用户不受锁定影响，应能正常登录"

    finally:
        for username in test_users:
            _delete_user(page, username)
