from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "pp02_1role"
_USER_NAME = "pp02_1user"
_PWD_0     = "Admin@11001"   # 初始密码
_PWD_1     = "Admin@22002"   # 第一次变更

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


def _set_history(page, value):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Password History")
    inp.fill(str(value))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_history(page):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Password History")
    inp.fill("1")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _create_role(page):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_ROLE_NAME)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="view", exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.wait_for_timeout(500)
    try:
        page.get_by_role("button", name="Yes, continue").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


def _create_user(page, username, password):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=_ROLE_NAME).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user_if_exists(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _admin_change_pwd(page, username, new_pwd) -> bool:
    """Admin changes user's password via Password Management. Returns True on success."""
    _nav_to_submenu(page, "Password Management")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").click()
    page.wait_for_timeout(500)
    page.get_by_label("Password", exact=True).fill(new_pwd)
    page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)
    changed = page.get_by_text("password changed", exact=False).is_visible()
    if not changed:
        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
            page.wait_for_timeout(300)
        except Exception:
            pass
    return changed


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


# 用例编号：TestCase_AcuHMI_007_03_case02_1
# 用例标题：设置密码历史记录为 1（无限制），重新设置密码为之前的密码，设置成功，登录成功
# 说明：Password History=1 表示"无限制"，可以复用之前的密码
# 测试步骤：
#   1. Password History = 1，点击 Save
#   2. 创建用户，管理员修改密码为 P1
#   3. 管理员将密码改回 P0（之前的密码）
#   4. 用 P0 登录系统
# 预期结果：
#   2. 配置成功
#   3. 设置成功（history=1 无限制，允许复用）
#   4. 登录成功
def test_TestCase_AcuHMI_007_03_case02_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    # 预清理：删除上次残留的同名用户/角色（崩溃中断会留孤儿，导致重名创建失败）。
    _delete_user_if_exists(page, _USER_NAME)
    _delete_role(page)

    _create_role(page)
    try:
        # Step 1: 设置 Password History = 1（无限制）
        _set_history(page, 1)

        # Step 2: 创建用户，修改密码到 P1
        _create_user(page, _USER_NAME, _PWD_0)
        assert _admin_change_pwd(page, _USER_NAME, _PWD_1), \
            "修改密码到 P1 应成功"

        # Step 3: 将密码改回 P0（history=1 = 无限制，应成功）
        assert _admin_change_pwd(page, _USER_NAME, _PWD_0), \
            "History=1（无限制）下，复用之前密码 P0 应成功"

        # Step 4: 用 P0 登录
        assert _can_login(browser, _USER_NAME, _PWD_0), \
            f"用 P0 应能登录用户 {_USER_NAME}"
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_history(page)
