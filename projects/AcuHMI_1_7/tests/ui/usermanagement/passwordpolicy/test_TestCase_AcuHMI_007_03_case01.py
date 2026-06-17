import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "pp01role"
_USER_NAME = "pp01user"
# 满足 Upper+Lower 策略的密码（包含大写+小写，无需数字或特殊字符）
_GOOD_PWD  = "TestAbcd"

_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]

_IDX_UPPER_LOWER     = 0
_IDX_NUMBERS_LETTERS = 1
_IDX_SPECIAL_CHARS   = 2


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


def _is_cb_checked(page, idx: int) -> bool:
    return page.evaluate(f"document.querySelectorAll('input[type=checkbox]')[{idx}].checked")


def _set_cb(page, idx: int, desired: bool):
    if _is_cb_checked(page, idx) != desired:
        page.locator(".el-checkbox__inner").nth(idx).click()
        page.wait_for_timeout(200)


def _set_policy(page, upper_lower: bool, numbers_letters: bool, special_chars: bool):
    _nav_to_submenu(page, "Password Policy")
    _set_cb(page, _IDX_UPPER_LOWER,      upper_lower)
    _set_cb(page, _IDX_NUMBERS_LETTERS,  numbers_letters)
    _set_cb(page, _IDX_SPECIAL_CHARS,    special_chars)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_default_policy(page):
    _set_policy(page, True, True, True)


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


def _delete_user_if_exists(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_03_case01
# 用例标题：配置密码策略，仅勾选大写和小写，创建用户密码包含大小写，可登录成功
# 测试步骤：
#   1. Password Policy 仅勾选 Upper and Lower Case，保存
#   2. 创建用户，密码为 "TestAbcd"（含大写+小写，无数字/特殊字符）
#   3. 新用户登录系统
# 预期结果：
#   1. 提示 "Password policy configuration saved"
#   2. 添加成功
#   3. 登录成功
def test_TestCase_AcuHMI_007_03_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    _create_role(page)
    try:
        _set_policy(page, upper_lower=True, numbers_letters=False, special_chars=False)

        _nav_to_submenu(page, "User Configuration")
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)
        page.get_by_label("Username", exact=True).fill(_USER_NAME)
        page.get_by_label("Password", exact=True).fill(_GOOD_PWD)
        page.get_by_label("Repeat Password", exact=True).fill(_GOOD_PWD)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name=_ROLE_NAME).click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        assert not page.get_by_label("Password", exact=True).is_visible(), \
            "Upper+Lower 策略下，包含大小写的密码创建用户应成功（对话框应关闭）"

        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(_USER_NAME)
            p.get_by_role("textbox", name="Enter Password").fill(_GOOD_PWD)
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
            assert "/#/login" not in p.url, \
                f"Upper+Lower 策略用户应能登录，当前 URL: {p.url}"
        finally:
            ctx.close()
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_default_policy(page)
