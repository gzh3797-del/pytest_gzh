import pytest
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


def _edit_user_role(page, username: str, new_role: str):
    """Open User Configuration edit dialog and change user's role."""
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    # Non-admin user rows have buttons: Lock(0), Edit(1), Delete(last)
    row.get_by_role("button").nth(1).click()
    page.wait_for_timeout(1000)
    # Update role - click the role dropdown (shows current role)
    page.locator(".el-form-item").filter(has_text="Role").locator(".el-select").click()
    page.wait_for_timeout(300)
    page.get_by_role("option", name=new_role).click()
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


def _login_verify(browser, username: str, password: str) -> bool:
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


# 用例编号：TestCase_AcuHMI_007_01_case05
# 用例标题：修改2个用户权限（角色互换），登录所有用户，查看用户权限
# 预置条件：服务启动正常，账号登录成功
# 测试步骤：
#   1. 创建 uc05user1（view）、uc05user2（admin），密码均为 pwd
#   2. 通过 User Configuration 编辑，将角色互换（user1: view→admin, user2: admin→view）
#   3. 用原密码登录两个用户（密码未变，均应成功）
#   4. 验证表格中两个用户的角色已正确更新
# 预期结果：
#   角色互换后，用户仍可用原密码登录；表中角色显示正确
# 备注：User Configuration 编辑对话框不支持密码修改（无 Password 字段），
#        密码修改需通过 Password Management 完成
def test_TestCase_AcuHMI_007_01_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    pwd = "Abc@11111"
    users = [
        ("uc05user1", "view", "admin"),  # (username, initial_role, new_role)
        ("uc05user2", "admin", "view"),
    ]

    try:
        for username, init_role, _ in users:
            _create_user(page, username, pwd, role=init_role)

        # Edit roles (User Configuration edit dialog supports role change only)
        for username, _, new_role in users:
            _edit_user_role(page, username, new_role)

        # Password is unchanged — users should still login with original password
        for username, _, _ in users:
            assert _login_verify(browser, username, pwd), \
                f"角色修改后，用户 {username} 应仍能用原密码登录"

        # Verify roles updated in table
        _nav_to_submenu(page, "User Configuration")
        for username, _, new_role in users:
            row = page.locator("tbody").get_by_role("row").filter(has_text=username)
            assert row.filter(has_text=new_role).count() > 0, \
                f"用户 {username} 的角色应已更新为 {new_role}"
    finally:
        for username, _, _ in users:
            _delete_user(page, username)
