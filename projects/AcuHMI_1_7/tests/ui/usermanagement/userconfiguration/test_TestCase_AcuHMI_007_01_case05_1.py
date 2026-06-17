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


def _edit_user_role(page, username: str, new_role: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").nth(1).click()  # Edit (Lock is first for non-admin)
    page.wait_for_timeout(1000)
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


# 用例编号：TestCase_AcuHMI_007_01_case05_1
# 用例标题：修改全部-1个用户权限，登录所有用户，查看该全部-1个用户权限
# 测试步骤：
#   1. 创建2个角色（uc051roleA=none, uc051roleB=view）
#   2. 创建5个用户，均绑定 roleA
#   3. 修改其中4个用户的角色为 roleB（保留1个仍为roleA）
#   4. 验证被修改的4个用户角色已更新（行显示 roleB），未修改的1个仍为 roleA
# 预期结果：
#   被修改用户角色均正确；未修改用户角色不变
def test_TestCase_AcuHMI_007_01_case05_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    role_a = "uc051roleA"
    role_b = "uc051roleB"
    pwd = "Abc@12345"

    to_change = [f"uc051u{i}" for i in range(1, 5)]   # 4 users to edit
    to_keep = ["uc051u5"]                               # 1 user unchanged
    all_users = to_change + to_keep

    try:
        _create_role(page, role_a, "none")
        _create_role(page, role_b, "view")

        for username in all_users:
            _create_user(page, username, pwd, role=role_a)

        # Edit 4 users to role_b (all-1)
        for username in to_change:
            _edit_user_role(page, username, role_b)

        # Verify role changes in the UI
        _nav_to_submenu(page, "User Configuration")
        for username in to_change:
            row = page.locator("tbody").get_by_role("row").filter(has_text=username)
            assert row.filter(has_text=role_b).count() > 0, \
                f"用户 {username} 的角色应已更新为 {role_b}"

        # Verify unchanged user still has role_a
        for username in to_keep:
            row = page.locator("tbody").get_by_role("row").filter(has_text=username)
            assert row.filter(has_text=role_a).count() > 0, \
                f"用户 {username} 的角色应仍为 {role_a}（未被修改）"
    finally:
        for username in all_users:
            _delete_user(page, username)
        _delete_role(page, role_a)
        _delete_role(page, role_b)
