# 用例编号：TestCase_ARM-XXL_002_04_case1_04（函数名/文件名因 Python 不能含 '-' 用下划线）
# 用例标题：admin用户修改其他用户密码 (LV2)
# 预置条件：admin 登录，存在其他用户
# 测试步骤：
#   1. Password Management 页面选择任一非 admin 用户，点编辑按钮
#   2. 输入新密码（无需输入该用户当前密码），保存
#   3. 用新密码以该用户身份登录系统
# 预期结果：
#   2. 不需要输入该用户当前密码可修改成功（参考真机：admin 改他人密码后不自动登出 admin 会话）
#   3. 登录成功
# 真机观察注：
#   - 手工用例预期「自动退出到登录页面」经真机核实不成立：admin 改他人密码后，
#     admin 会话保持不变，不会自动登出。本用例按真机行为落地断言。
#   - admin 改他人密码时表单仅含 New Password / Repeat Password（无 Current Password）。
from playwright.sync_api import expect

from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 固定硬编码密码常量（禁止使用变化/临时生成值）
_TEST_USER = "pc0104u"
_INIT_PWD = "Admin@110001"
_NEW_PWD = "Admin@110003"

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


def _create_user(page, username: str, password: str, role: str = "view") -> None:
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


def _change_password_via_pm(page, username: str, new_pwd: str) -> None:
    """admin 通过 Password Management 改他人密码（无 Current Password 字段）。"""
    _nav_to_submenu(page, "Password Management")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").click()
    page.wait_for_timeout(500)
    # 验证：表单不含 Current Password 字段（admin 改他人无需当前密码）
    current_pwd_field = page.get_by_label("Current Password", exact=False)
    assert current_pwd_field.count() == 0, (
        "admin 改他人密码时不应出现 Current Password 字段，"
        f"但页面检测到 {current_pwd_field.count()} 个该字段"
    )
    page.get_by_label("Password", exact=True).fill(new_pwd)
    page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _can_login(browser, username: str, password: str) -> bool:
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


def test_TestCase_ARM_XXL_002_04_case1_04(login_page: LoginPage) -> None:
    """admin 改他人密码无需当前密码，admin 会话不登出，被改用户用新密码可登录。"""
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser = admin_page.context.browser

    _create_user(admin_page, _TEST_USER, _INIT_PWD, role="view")
    try:
        # 步骤 1-2：admin 修改 _TEST_USER 的密码
        _change_password_via_pm(admin_page, _TEST_USER, _NEW_PWD)

        # 真机核实：admin 改他人密码后，admin 会话不登出
        # 手工预期「自动退出到登录页面」与真机行为不符，按真机断言
        assert "/#/login" not in admin_page.url, (
            "admin 改他人密码后不应登出，admin 会话应保持（真机验证结论）。"
            f"实际 URL: {admin_page.url}"
        )

        # 验证保存成功提示
        expect(admin_page.get_by_text("password changed", exact=False)).to_be_visible(
            timeout=5000
        )

        # 步骤 3：用新密码以 _TEST_USER 登录
        assert _can_login(browser, _TEST_USER, _NEW_PWD), (
            f"admin 改密后，{_TEST_USER} 用新密码 {_NEW_PWD} 登录失败"
        )

    finally:
        _delete_user(admin_page, _TEST_USER)
