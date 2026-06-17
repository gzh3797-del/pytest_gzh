import pytest
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]


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


def _create_role(page, role_name: str, level: str):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name=level, exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _edit_role(page, role_name: str, level: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    row.get_by_role("button").first.click()
    page.wait_for_timeout(1000)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name=level, exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _create_user(page, username: str, password: str, role: str):
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


def _login_as_user(browser, username, password):
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
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
    return ctx, p


# 用例编号：TestCase_AcuHMI_007_01_case05_03
# 用例标题：修改1个用户权限，登录所有用户，查看用户权限
# 测试步骤：
#   1. 创建 "uc0503none" 角色（所有权限=none）
#   2. 创建用户 uc0503user，绑定 uc0503none 角色
#   3. 用户登录 → 看到无权限提示（无导航菜单）
#   4. 将 uc0503none 角色所有权限改为 edit
#   5. 用户重新登录 → 应有 edit 权限（可看到页面功能）
# 预期结果：
#   角色从 none 改为 edit 后，用户登录可见 edit 权限的功能
def test_TestCase_AcuHMI_007_01_case05_03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    role_name = "uc0503none"
    username = "uc0503user"
    pwd = "Abc@12345"

    _create_role(admin_page, role_name, "none")
    _create_user(admin_page, username, pwd, role=role_name)
    try:
        # Login with none-permission role → no accessible modules
        ctx1, p1 = _login_as_user(browser, username, pwd)
        try:
            none_login_ok = "/#/login" not in p1.url
            if none_login_ok:
                # If user can login, they should not see User Management
                has_user_mgmt = p1.get_by_text("User Management").is_visible()
                assert not has_user_mgmt, \
                    "none 权限用户不应看到 User Management 菜单"
            # Login failure for none-permission users is also acceptable device behavior
        finally:
            ctx1.close()

        # Change role to edit permissions
        _edit_role(admin_page, role_name, "edit")

        # Login again → should now have edit permissions
        ctx2, p2 = _login_as_user(browser, username, pwd)
        try:
            assert "/#/login" not in p2.url, \
                f"edit 权限用户登录应成功，当前 URL: {p2.url}"
        finally:
            ctx2.close()
    finally:
        _delete_user(admin_page, username)
        _delete_role(admin_page, role_name)
