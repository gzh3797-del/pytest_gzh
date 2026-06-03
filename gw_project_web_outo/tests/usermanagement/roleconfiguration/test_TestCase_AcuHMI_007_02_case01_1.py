import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME = "rc01_1view"
_USER_NAME = "rcuser01_1"
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


def _create_role_all_view(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="view", exact=True).click()
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


# 用例编号：TestCase_AcuHMI_007_02_case01_1
# 用例标题：添加角色，角色权限均为视图，创建该用户，登录后查看该用户权限
# 预置条件：管理权限登录AcuHMI网页
# 测试步骤：
#   1. Role Configuration -> Add Role，所有权限设置 view，点击 Save
#   2. User Configuration -> Add User，角色设置为该自定义角色
#   3. 新用户登录系统
# 预期结果：
#   1. 保存成功
#   2. 创建成功
#   3. 登录成功，系统权限为视图（User Config 中无 Add User 按钮）
def test_TestCase_AcuHMI_007_02_case01_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    _create_role_all_view(admin_page, _ROLE_NAME)
    _create_user(admin_page, _USER_NAME, _INIT_PWD, role=_ROLE_NAME)
    try:
        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(_USER_NAME)
            p.get_by_role("textbox", name="Enter Password").fill(_INIT_PWD)
            p.get_by_role("button", name="Sign In").click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(1000)

            # 处理 EULA
            try:
                p.get_by_role("button", name="Accept").click(timeout=3000)
                p.wait_for_load_state("networkidle")
                p.wait_for_timeout(500)
            except Exception:
                pass

            # 处理默认密码提示
            try:
                p.get_by_role("button", name="Cancel").click(timeout=2000)
            except Exception:
                pass

            # 登录成功
            assert "/#/login" not in p.url, \
                f"all-view 角色用户登录应成功，当前 URL: {p.url}"

            # 导航到 User Configuration，view 角色无 Add User 按钮
            _nav_to_submenu(p, "User Configuration")
            add_user_btn = p.get_by_role("button", name="Add User")
            assert not add_user_btn.is_visible(), \
                "all-view 角色用户在 User Configuration 中不应有 Add User 按钮（视图权限）"
        finally:
            ctx.close()
    finally:
        _delete_user(admin_page, _USER_NAME)
        _delete_role(admin_page, _ROLE_NAME)
