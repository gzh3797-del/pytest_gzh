import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_USER = "pwdrst02"
_INIT_PWD  = "Admin@110001"
_NEW_PWD   = "Admin@Reset02"


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


def _create_user(page, username: str, password: str, role: str = "view"):
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
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _change_password_via_pm(page, username: str, new_pwd: str):
    _nav_to_submenu(page, "Password Management")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").click()
    page.wait_for_timeout(500)
    page.get_by_label("Password", exact=True).fill(new_pwd)
    page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _can_login(browser, username: str, password: str) -> bool:
    ctx = browser.new_context(ignore_https_errors=True)
    p   = ctx.new_page()
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


# 用例编号：TestCase_ARM_XXL_002_04_case12_02
# 用例标题：非admin用户点击Forgot password弹框提示联系管理员，管理员修改密码后登录成功
# 预置条件：管理权限登录 AcuHMI 网页，存在 view 权限用户
# 测试步骤：
#   1. 登录页面输入用户名，点击 Forgot password
#   2. 弹框显示 "Please contact your administrator for assistance"
#   3. 管理员通过 Password Management 修改该用户密码
#   4. 用新密码登录
# 预期结果：
#   2. 弹框显示 "Please contact your administrator for assistance"
#   4. 登录成功
def test_TestCase_ARM_XXL_002_04_case12_02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser    = admin_page.context.browser

    _create_user(admin_page, _TEST_USER, _INIT_PWD, role="view")
    try:
        # 步骤 1：在登录页面输入用户名，点击 Forgot password
        ctx = browser.new_context(ignore_https_errors=True)
        p   = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(_TEST_USER)
            p.get_by_text("Forgot password", exact=False).click()
            p.wait_for_timeout(1000)

            # 步骤 2：验证弹框提示联系管理员
            expect(
                p.get_by_text("Please contact your administrator for assistance", exact=False)
            ).to_be_visible(timeout=5000)

            # 关闭弹框（尝试常见按钮名称）
            for btn_name in ["OK", "Ok", "Close", "确认", "关闭"]:
                try:
                    p.get_by_role("button", name=btn_name).click(timeout=2000)
                    p.wait_for_timeout(300)
                    break
                except Exception:
                    pass
        finally:
            ctx.close()

        # 步骤 3：管理员修改该用户密码
        _change_password_via_pm(admin_page, _TEST_USER, _NEW_PWD)
        expect(admin_page.get_by_text("password changed", exact=False)).to_be_visible(
            timeout=5000
        )

        # 步骤 4：用新密码登录验证
        assert _can_login(browser, _TEST_USER, _NEW_PWD), \
            f"管理员改密后，用户 {_TEST_USER} 用新密码登录失败"
    finally:
        _delete_user(admin_page, _TEST_USER)
