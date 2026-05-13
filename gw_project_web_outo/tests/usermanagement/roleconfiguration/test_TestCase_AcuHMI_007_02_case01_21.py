import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME = "rc01_21fwe"
_USER_NAME = "rcuser01_21"
_INIT_PWD  = "Admin@110001"

# Firmware Update=edit, all others=none
_PERM_MAP = {
    "User":             "none",
    "Device":           "none",
    "Data Log":         "none",
    "System Settings":  "none",
    "Protocol":         "none",
    "Alarm Log":        "none",
    "Maintenance":      "none",
    "Diagnostics":      "none",
    "Firmware Update":  "edit",
}


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


def _create_role(page, role_name: str, perm_map: dict):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl, val in perm_map.items():
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name=val, exact=True).click()
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


# 用例编号：TestCase_AcuHMI_007_02_case01_21
# 用例标题：添加角色，角色权限为固件更新-编辑，其余均为无，创建该用户，登录后验证权限
# 测试步骤：
#   1. Role Configuration -> Add Role，Firmware Update=edit，其余均为 none
#   2. User Configuration -> Add User，角色设置为该自定义角色
#   3. 新用户登录系统，落地到 Firmware Update 页面
#   4. 确认页面有 Browse 按钮（edit 权限）
#   5. 点击顶部导航 "Devices"
# 预期结果：
#   3. 登录成功，落地于 /#/firmwareUpdate 页面
#   4. Browse 按钮可见
#   5. 弹出 "No Any Permissions" 提示（用户无权限访问 Devices 区域）
def test_TestCase_AcuHMI_007_02_case01_21(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    _create_role(admin_page, _ROLE_NAME, _PERM_MAP)
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

            # Step 3: 登录成功，落地于 Firmware Update 页面
            assert "/#/login" not in p.url, \
                f"Firmware Update=edit 用户登录应成功，当前 URL: {p.url}"
            assert "firmwareUpdate" in p.url, \
                f"Firmware Update=edit 用户应落地于 firmwareUpdate 页面，当前 URL: {p.url}"

            # Step 4: edit 权限用户的 Firmware Update 页面应有 Browse 按钮
            browse_btn = p.get_by_role("button", name="Browse")
            assert browse_btn.is_visible(), \
                "Firmware Update=edit 用户在 firmwareUpdate 页面应有 Browse 按钮"

            # Step 5: 点击 "Devices" → 应弹出 "No Any Permissions" toast
            p.locator("header span").filter(has_text="Devices").first.click()
            toast = p.locator(".el-message__content").filter(has_text="No Any Permissions")
            toast.wait_for(state="visible", timeout=5000)
            assert toast.is_visible(), \
                "点击 Devices 后应弹出 'No Any Permissions' 提示（用户无权限访问该区域）"
        finally:
            ctx.close()
    finally:
        _delete_user(admin_page, _USER_NAME)
        _delete_role(admin_page, _ROLE_NAME)
