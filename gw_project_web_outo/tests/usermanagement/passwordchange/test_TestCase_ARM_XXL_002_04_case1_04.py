import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_USER = "pwdchg04"
_INIT_PWD  = "Admin@110001"
_NEW_PWD   = "Admin@NewPwd1"


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


def _create_user(page, username: str, password: str):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name="view").click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


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
        for btn in ["Accept", "I Accept", "确认"]:
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


# ── 用例编号：TestCase_ARM_XXL_002_04_case1_04
# 用例标题：admin 用户修改其他用户密码
# 预置条件：
#   1. admin 用户已登录系统
#   2. 存在其他非 admin 用户
# 测试步骤：
#   1. Password Management 页面选择非 admin 用户，点击编辑按钮
#   2. 输入新密码，点击保存（无需输入该用户当前密码）
#   3. 用新密码登录系统
# 预期结果：
#   2. 不需要输入该用户当前密码，修改成功（提示 "User password changed"）
#   3. 新密码登录成功
def test_TestCase_ARM_XXL_002_04_case1_04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page    = login_page.page
    browser = page.context.browser

    _create_user(page, _TEST_USER, _INIT_PWD)
    try:
        # ── 步骤 1：进入 Password Management，找到非 admin 用户并点击编辑 ────
        _nav_to_submenu(page, "Password Management")
        row = page.locator("tbody").get_by_role("row").filter(has_text=_TEST_USER)
        row.get_by_role("button").click()
        page.wait_for_timeout(500)

        # ── 步骤 2：验证无"Current Password"字段，只输入新密码，点击 Save ────
        # admin 修改其他用户密码时，表单中不应有 Current Password 字段
        current_pwd_visible = page.get_by_label("Current Password", exact=True).is_visible()
        assert not current_pwd_visible, \
            "admin 修改其他用户密码时，不应出现 Current Password 字段"

        page.get_by_label("Password", exact=True).fill(_NEW_PWD)
        page.get_by_label("Repeat Password", exact=True).fill(_NEW_PWD)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        expect(page.get_by_text("password changed", exact=False)).to_be_visible(
            timeout=5000
        )

        # ── 步骤 3：用新密码登录验证 ─────────────────────────────────────────
        assert _can_login(browser, _TEST_USER, _NEW_PWD), \
            f"admin 改密后，用新密码登录用户 {_TEST_USER} 失败"
    finally:
        _delete_user(page, _TEST_USER)