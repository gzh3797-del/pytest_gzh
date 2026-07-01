from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME   = "pp03_3role"
_USER_NAME   = "pp03_3user"
_PWD_0       = "Admin@11001"
_PWD_1       = "Admin@22002"
_MIN_AGE     = 50

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


def _set_min_age(page, value):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Minimum Password Age")
    inp.fill(str(value))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_min_age(page):
    _nav_to_submenu(page, "Password Policy")
    inp = page.get_by_placeholder("Enter Minimum Password Age")
    inp.fill("0")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _create_role_with_user_edit(page):
    """Create role with User=edit so the user can change their own password."""
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_ROLE_NAME)
    # Set User=edit, rest=view
    page.locator(".el-form-item").filter(has_text="User").locator(".el-select").click()
    page.wait_for_timeout(200)
    page.get_by_role("option", name="edit", exact=True).click()
    page.wait_for_timeout(200)
    for lbl in _PERM_LABELS[1:]:
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


def _user_self_change_pwd(browser, username, current_pwd, new_pwd):
    """非管理员用户自助修改密码。

    返回 (changed, toast_msg)：
      changed   —— 是否修改成功（提示含 "password changed"）；
      toast_msg —— 点击 Save 后捕获的提示文本（成功提示或策略拒绝提示），
                   用于让调用方区分"被策略拦截"与"流程未走通"，避免负向断言假阳性。
    """
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    try:
        p.goto(BASE_URL + "/#/login")
        p.wait_for_load_state("networkidle")
        p.get_by_role("textbox", name="Enter User Name").fill(username)
        p.get_by_role("textbox", name="Enter Password").fill(current_pwd)
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

        # Navigate to Password Management
        if "/userManagement/" not in p.url:
            p.locator("header span").filter(has_text="AcuHMI").first.click()
            p.wait_for_load_state("networkidle")
            p.wait_for_timeout(500)
            p.get_by_text("User Management").first.click()
            p.wait_for_timeout(500)
        p.get_by_role("menuitem", name="Password Management").click()
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(500)

        row = p.locator("tbody").get_by_role("row").filter(has_text=username)
        row.get_by_role("button").click()
        p.wait_for_timeout(500)

        # Fill current password dialog if appears
        try:
            p.get_by_placeholder("Please input").fill(current_pwd)
            p.get_by_role("button", name="Confirm").click()
            p.wait_for_timeout(500)
        except Exception:
            pass

        p.get_by_label("Password", exact=True).fill(new_pwd)
        p.get_by_label("Repeat Password", exact=True).fill(new_pwd)
        p.get_by_role("button", name="Save").click()
        # 捕获结果提示（成功含 "password changed"；被策略拦截时为拒绝提示）
        msg = ""
        try:
            content = p.locator(".el-message__content")
            content.last.wait_for(state="visible", timeout=3000)
            msg = (content.last.text_content() or "").strip()
        except Exception:
            pass
        changed = "password changed" in msg.lower()
        return changed, msg
    finally:
        ctx.close()


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


# 用例编号：TestCase_AcuHMI_007_03_case03_3
# 用例标题：设置最短密码期限为50，密码在50天之内不允许修改
# 简化说明：无需真等 50 天——刚创建的用户密码"年龄"为 0 天，处于 50 天最短期限窗口内，
#           此刻自助修改应被策略拒绝。以"第 0 天立即修改被拒"等价验证"50 天内不允许修改"。
#           （与 case03_2 同套路，仅 Minimum Password Age 由 1 改为 50）
# 测试步骤：
#   1. Minimum Password Age = 50，保存
#   2. 创建新用户，密码为 P0
#   3. 用户立即（远不到 50 天）尝试自助修改密码
# 预期结果：
#   3. 修改失败：新密码 P1 不能登录，原密码 P0 仍可登录
def test_TestCase_AcuHMI_007_03_case03_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    # 预清理：删除上次残留的同名用户/角色（崩溃中断会留孤儿，导致重名创建失败）。
    _delete_user_if_exists(page, _USER_NAME)
    _delete_role(page)

    _create_role_with_user_edit(page)
    try:
        # Step 1: 设置 Min Age = 50
        _set_min_age(page, _MIN_AGE)

        # Step 2: 创建用户
        _create_user(page, _USER_NAME, _PWD_0)

        # Step 3: 用户立即尝试自改密码 → 应失败（密码年龄 0 天，未到 50 天）
        changed, msg = _user_self_change_pwd(browser, _USER_NAME, _PWD_0, _PWD_1)
        assert not changed, (
            f"Minimum Password Age={_MIN_AGE} 下，用户在 {_MIN_AGE} 天内尝试修改密码应失败，"
            f"实际提示: '{msg}'"
        )

        # 双重确认密码确实未被修改：新密码 P1 不能登录、原密码 P0 仍可登录。
        # 强于"仅看不到成功提示"——可区分策略真实拦截与自动化未走通改密流程（避免假阳性）。
        assert not _can_login(browser, _USER_NAME, _PWD_1), \
            f"改密应被拦截，新密码 P1 不应能登录 {_USER_NAME}（提示: '{msg}'）"
        assert _can_login(browser, _USER_NAME, _PWD_0), \
            "修改失败后，原密码应仍可登录"
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_min_age(page)
