import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_P8  = "Admin@12"
_P20 = "Admin@12" + "x" * 12
_P40 = "Admin@12" + "x" * 32
_INIT_PWD  = "Admin@110001"
_TEST_USER = "pwdcfg01"


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


# ── 用例编号：TestCase_AcuHMI_007_04_case01
# 用例标题：修改密码，密码长度为8、20、40，保存配置成功，新密码可登录成功
# 预置条件：管理权限登录 AcuHMI 网页
# 测试步骤：
#   1. Password Management 页面选择任一用户，点击编辑按钮
#   2. 修改密码匹配当前密码规则，密码长度分别为 8、20、40
#   3. 验证保存成功（提示 "User password changed"）
#   4. 用新密码登录系统，验证登录成功
# 预期结果：
#   2. 保存成功，提示 "User password changed"
#   4. 新密码登录成功
def test_TestCase_AcuHMI_007_04_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page    = login_page.page
    browser = page.context.browser

    _create_user(page, _TEST_USER, _INIT_PWD)
    try:
        for pwd in [_P8, _P20, _P40]:
            _change_password_via_pm(page, _TEST_USER, pwd)
            expect(page.get_by_text("password changed", exact=False)).to_be_visible(
                timeout=5000
            )
            assert _can_login(browser, _TEST_USER, pwd), \
                f"新密码（{len(pwd)}位）登录失败：用户 {_TEST_USER}"
    finally:
        _delete_user(page, _TEST_USER)
