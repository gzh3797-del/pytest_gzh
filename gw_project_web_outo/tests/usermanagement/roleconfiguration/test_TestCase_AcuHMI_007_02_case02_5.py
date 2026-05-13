import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME = "rc02_5edit2view"
_USER_NAME = "rcuser02_5"
_INIT_PWD  = "Admin@110001"

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


# 用例编号：TestCase_AcuHMI_007_02_case02_5
# 用例标题：编辑角色，由角色权限全编辑改为全视图，创建该用户，登录后查看该用户权限
# 测试步骤：
#   1. test3 用户角色权限为全编辑，登录该用户查看权限
#   2. 修改该用户角色权限为全视图，登录该用户查看权限
# 预期结果：
#   1. 该用户权限均为编辑（User Configuration 有 Add User 按钮）
#   2. 该用户权限均为只读（User Configuration 无 Add User 按钮）
def test_TestCase_AcuHMI_007_02_case02_5(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    _create_role(admin_page, _ROLE_NAME, "edit")
    _create_user(admin_page, _USER_NAME, _INIT_PWD, role=_ROLE_NAME)
    try:
        # Step 1: 全编辑 → User Config 有 Add User
        ctx1, p1 = _login_as_user(browser, _USER_NAME, _INIT_PWD)
        try:
            assert "/#/login" not in p1.url, \
                f"全编辑角色用户登录应成功，当前 URL: {p1.url}"
            _nav_to_submenu(p1, "User Configuration")
            expect(p1.get_by_role("button", name="Add User")).to_be_visible(timeout=5000)
        finally:
            ctx1.close()

        # 编辑角色为全视图
        _edit_role(admin_page, _ROLE_NAME, "view")

        # Step 2: 全视图 → User Config 无 Add User
        ctx2, p2 = _login_as_user(browser, _USER_NAME, _INIT_PWD)
        try:
            assert "/#/login" not in p2.url, \
                f"全视图角色用户登录应成功，当前 URL: {p2.url}"
            _nav_to_submenu(p2, "User Configuration")
            assert not p2.get_by_role("button", name="Add User").is_visible(), \
                "全视图角色用户在 User Configuration 中不应有 Add User 按钮"
        finally:
            ctx2.close()
    finally:
        _delete_user(admin_page, _USER_NAME)
        _delete_role(admin_page, _ROLE_NAME)
