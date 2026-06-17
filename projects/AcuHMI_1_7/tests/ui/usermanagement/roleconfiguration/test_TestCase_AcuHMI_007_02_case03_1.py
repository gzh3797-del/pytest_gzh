import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_A = "rc03_1a"
_ROLE_B = "rc03_1b"
_USER_A = "rcuser03_1a"
_USER_B = "rcuser03_1b"
_INIT_PWD = "Admin@110001"

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


def _try_delete_role_expect_fail(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(1000)
    # verify role still in table
    _nav_to_submenu(page, "Role Configuration")
    row_after = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    assert row_after.count() > 0, \
        f"角色 '{role_name}' 已有用户绑定，删除应失败，但角色已从表中消失"


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


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


# 用例编号：TestCase_AcuHMI_007_02_case03_1
# 用例标题：删除2个角色，登录拥有该角色的用户，查看用户登录是否正常
# 测试步骤：
#   1. Role Configuration，选择2个已有用户使用的角色，逐一点击删除
#   2. User Configuration 先删除用户1和2，再删除2个角色
# 预期结果：
#   1. 删除失败，提示已有用户使用该角色，无法删除（两个角色均如此）
def test_TestCase_AcuHMI_007_02_case03_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page

    _create_role(admin_page, _ROLE_A)
    _create_role(admin_page, _ROLE_B)
    _create_user(admin_page, _USER_A, _INIT_PWD, role=_ROLE_A)
    _create_user(admin_page, _USER_B, _INIT_PWD, role=_ROLE_B)
    try:
        # Step 1: 两个角色均有用户，删除均应失败
        _try_delete_role_expect_fail(admin_page, _ROLE_A)
        _try_delete_role_expect_fail(admin_page, _ROLE_B)
    finally:
        _delete_user(admin_page, _USER_A)
        _delete_user(admin_page, _USER_B)
        _delete_role(admin_page, _ROLE_A)
        _delete_role(admin_page, _ROLE_B)
