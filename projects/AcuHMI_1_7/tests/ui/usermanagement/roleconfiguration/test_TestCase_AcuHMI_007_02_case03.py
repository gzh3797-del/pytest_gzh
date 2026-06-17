import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "rc03_a"
_USER_NAME = "rcuser03_a"
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


def _create_role(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="none", exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.wait_for_timeout(500)
    try:
        page.get_by_role("button", name="Yes, continue").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


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


# 用例编号：TestCase_AcuHMI_007_02_case03
# 用例标题：删除1个角色，登录拥有该角色的用户，查看用户登录是否正常
# 测试步骤：
#   1. Role Configuration -> Action，选择已有用户使用的角色，点击删除
#   2. 先删除使用该角色的用户，再删除角色
# 预期结果：
#   1. 删除失败，提示已有用户使用该角色，无法删除
#   2. 删除角色成功
def test_TestCase_AcuHMI_007_02_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page

    _create_role(admin_page, _ROLE_NAME)
    _create_user(admin_page, _USER_NAME, _INIT_PWD, role=_ROLE_NAME)
    try:
        # Step 1: 尝试删除已有用户的角色 → 应失败（角色仍在表中）
        _nav_to_submenu(admin_page, "Role Configuration")
        row = admin_page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
        row.get_by_role("button").last.click()
        admin_page.wait_for_timeout(500)
        # 出现确认对话框，点击确认
        admin_page.get_by_role("button", name="Yes, continue").click()
        admin_page.wait_for_timeout(1000)
        # 重新加载 Role Configuration 检查角色是否仍然存在
        _nav_to_submenu(admin_page, "Role Configuration")
        row_after = admin_page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
        assert row_after.count() > 0, \
            f"角色 '{_ROLE_NAME}' 已有用户绑定，删除应失败，但角色已从表中消失"

        # Step 2: 先删除用户，再删除角色 → 角色应成功删除
        _delete_user(admin_page, _USER_NAME)
        _nav_to_submenu(admin_page, "Role Configuration")
        row2 = admin_page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
        row2.get_by_role("button").last.click()
        admin_page.wait_for_timeout(500)
        admin_page.get_by_role("button", name="Yes, continue").click()
        admin_page.wait_for_timeout(1000)
        _nav_to_submenu(admin_page, "Role Configuration")
        row_gone = admin_page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
        assert row_gone.count() == 0, \
            f"删除用户后再删除角色应成功，但角色 '{_ROLE_NAME}' 仍在表中"
    finally:
        _delete_user(admin_page, _USER_NAME)
        _delete_role(admin_page, _ROLE_NAME)
