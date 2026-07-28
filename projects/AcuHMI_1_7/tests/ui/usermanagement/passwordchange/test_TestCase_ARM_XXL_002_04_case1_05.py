# 用例编号：TestCase_ARM-XXL_002_04_case1_05（函数名/文件名因 Python 不能含 '-' 用下划线）
# 用例标题：非admin用户修改自己的密码 (LV2)
# 预置条件：管理权限登录网页
# 测试步骤：
#   1. 新建 edit 和 view 权限的角色，再用该角色创建用户 pc0105u1
#   2. pc0105u1 登录系统
#   3. Password Management 页面选择自己（pc0105u1）点编辑
#   4. 修改密码，确认是否需要输入当前用户密码
# 预期结果：
#   2. 登录成功
#   4. 修改成功，需要输入当前用户密码
# 真机观察注：
#   - 进 edit 页后立即弹出 .password-verify-dialog（标题 "Current User Password"），
#     内含 input[type=password][placeholder="Please input"] + Cancel/Confirm 按钮，
#     无 "Current Password" label。需先在弹窗里输入当前密码并点 Confirm，才能填编辑表单。
#   - 编辑表单字段：Username（只读）+ Password + Repeat Password（无 New Password 字段）。
#   - 非 admin 改自己密码后【不自动登出】（与 admin 改自身密码后自动登出不同），
#     跳转到 passwordManagement list 页并显示 toast "User <name> password changed"。
#   - 手工预期"修改成功，需要输入当前用户密码"确认成立；"自动退出到登录页面"与真机不符。
from playwright.sync_api import expect

from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 固定硬编码密码常量（禁止使用变化/临时生成值）
_ROLE_NAME = "pc0105r"
_USER1 = "pc0105u1"
_INIT_PWD = "Admin@110001"
_NEW_PWD = "Admin@110003"

# 角色权限：User=edit 保证该用户能进 Password Management；其余给 view/none 不干扰导航
_PERM_MAP = {
    "User": "edit",
    "Device": "view",
    "Data Log": "none",
    "System Settings": "none",
    "Protocol": "none",
    "Alarm Log": "view",
    "Maintenance": "none",
    "Diagnostics": "none",
    "Firmware Update": "none",
}

# ---------- helpers ----------


def _nav_to_submenu(page, submenu: str) -> None:
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _create_role(page, role_name: str, perm_map: dict) -> None:
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


def _delete_role(page, role_name: str) -> None:
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _create_user(page, username: str, password: str, role: str) -> None:
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


def _delete_user(page, username: str) -> None:
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_as(browser, username: str, password: str):
    """以指定账号登录，返回 (page, context)；调用方负责关闭 context。"""
    ctx = browser.new_context(ignore_https_errors=True)
    page = ctx.new_page()
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.get_by_placeholder("Enter User Name").fill(username)
    page.get_by_placeholder("Enter Password").fill(password)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    for btn in ["Accept", "I Accept"]:
        try:
            page.get_by_role("button", name=btn).click(timeout=2000)
            page.wait_for_load_state("networkidle")
        except Exception:
            pass
    try:
        page.get_by_role("button", name="Cancel").click(timeout=2000)
    except Exception:
        pass
    return page, ctx


def _can_login(browser, username: str, password: str) -> bool:
    """用独立 context 验证账号能否登录成功。"""
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    try:
        p.goto(BASE_URL + "/#/login")
        p.wait_for_load_state("networkidle")
        p.get_by_placeholder("Enter User Name").fill(username)
        p.get_by_placeholder("Enter Password").fill(password)
        p.get_by_role("button", name="Sign In").click()
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(1500)
        for btn in ["Accept", "I Accept"]:
            try:
                p.get_by_role("button", name=btn).click(timeout=2000)
                p.wait_for_load_state("networkidle")
            except Exception:
                pass
        try:
            p.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
        return "/#/login" not in p.url
    finally:
        ctx.close()


# ---------- test ----------


def test_TestCase_ARM_XXL_002_04_case1_05(login_page: LoginPage) -> None:
    """非admin用户（User=edit）修改自己密码：验证弹窗 + 新密码登录成功。

    真机结论：
    - 进 edit 页后立即弹出 .password-verify-dialog（标题 "Current User Password"），
      内含 input[type=password][placeholder='Please input']，无 Current Password label。
    - 弹窗先 Confirm，再填表单 Password/Repeat Password → Save。
    - 改密后不自动登出，停留在 passwordManagement list 页，显示 toast。
    - 用新密码独立 context 登录验证成功。
    """
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    # 步骤 1：新建角色和用户
    _create_role(admin_page, _ROLE_NAME, _PERM_MAP)
    _create_user(admin_page, _USER1, _INIT_PWD, role=_ROLE_NAME)

    user1_ctx = None
    try:
        # 步骤 2：pc0105u1 登录
        user1_page, user1_ctx = _login_as(browser, _USER1, _INIT_PWD)

        assert "/#/login" not in user1_page.url, (
            f"{_USER1} 登录应成功，实际 URL: {user1_page.url}"
        )

        # 步骤 3：导航到 Password Management
        # 真机确认：User=edit 的非admin用户可以访问 Password Management 菜单
        user1_page.locator("header span").filter(has_text="AcuHMI").first.click()
        user1_page.wait_for_load_state("networkidle")
        user1_page.wait_for_timeout(500)
        user1_page.get_by_text("User Management").first.click()
        user1_page.wait_for_timeout(500)
        user1_page.get_by_role("menuitem", name="Password Management").click()
        user1_page.wait_for_load_state("networkidle")
        user1_page.wait_for_timeout(1000)

        assert "passwordManagement" in user1_page.url, (
            f"{_USER1}（User=edit）应能访问 Password Management，"
            f"实际 URL: {user1_page.url}"
        )

        # 步骤 4：点击自身（pc0105u1）行的编辑按钮
        self_row = user1_page.locator("tbody").get_by_role("row").filter(has_text=_USER1)
        assert self_row.count() > 0, (
            f"Password Management 列表中未找到 {_USER1} 行"
        )
        self_row.get_by_role("button").click()
        user1_page.wait_for_timeout(1500)

        # 验证：弹出 .password-verify-dialog（真机结论：非admin改自己密码也弹此弹窗）
        # 弹窗内含 input[type=password][placeholder="Please input"] + Cancel/Confirm
        verify_dlg = user1_page.locator(".password-verify-dialog")
        assert verify_dlg.count() > 0, (
            "非 admin 修改自己密码时应弹出 .password-verify-dialog 验证弹窗，"
            "但页面上未检测到（手工预期 Current Password 字段实为弹窗形式，非表单 label）"
        )

        # 在弹窗内输入当前密码并 Confirm
        verify_input = verify_dlg.locator("input[type=password]")
        verify_input.fill(_INIT_PWD)
        verify_dlg.get_by_role("button", name="Confirm").click()
        user1_page.wait_for_timeout(1000)

        # 填写新密码表单（字段：Password / Repeat Password，无 New Password label）
        user1_page.get_by_label("Password", exact=True).fill(_NEW_PWD)
        user1_page.get_by_label("Repeat Password", exact=True).fill(_NEW_PWD)
        user1_page.get_by_role("button", name="Save").click()
        user1_page.wait_for_timeout(2000)

        # 改密成功断言：
        # 真机观察：非admin 改自己密码后【不自动登出】，跳转到 list 页并显示 toast
        # "User <name> password changed"（与 admin 改自身密码后自动登出不同）
        user1_page.wait_for_load_state("networkidle")
        user1_page.wait_for_timeout(1000)
        assert "passwordManagement" in user1_page.url, (
            f"{_USER1} 改密后应停留在 passwordManagement 页面（list），"
            f"实际 URL: {user1_page.url}"
        )
        # 验证成功 toast
        expect(user1_page.locator(".el-message").filter(
            has_text="password changed"
        )).to_be_visible(timeout=5000)

        # 步骤 5：用新密码登录成功（当前 session 未登出，用独立 context 验证）
        assert _can_login(browser, _USER1, _NEW_PWD), (
            f"{_USER1} 改密后用新密码 {_NEW_PWD} 登录失败"
        )

    finally:
        if user1_ctx is not None:
            user1_ctx.close()
        _delete_user(admin_page, _USER1)
        _delete_role(admin_page, _ROLE_NAME)
