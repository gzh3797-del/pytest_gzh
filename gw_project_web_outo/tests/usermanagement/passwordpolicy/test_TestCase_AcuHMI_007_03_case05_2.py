import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME = "pp05_2role"
_USER_NAME = "pp05_2user"
_SHORT_PWD = "Ac@123"        # 6 chars — too short for min=40
# 40 chars: "Admin@110001" × 3 + "1234" = 36 + 4 = 40
_GOOD_PWD  = "Admin@110001Admin@110001Admin@1100011234"

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


def _set_min_length(page, value):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Minimum Password Length").fill(str(value))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_length(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Minimum Password Length").fill("8")
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


def _try_create_user(page, username, password):
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
    page.wait_for_timeout(1500)
    has_error = (page.locator(".el-form-item__error").count() > 0 or
                 page.locator(".el-message--error").count() > 0)
    if has_error:
        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
            page.wait_for_timeout(500)
        except Exception:
            pass
        return False
    return True


def _delete_user_if_exists(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
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


# 用例编号：TestCase_AcuHMI_007_03_case05_2
# 用例标题：设置最小密码长度为 40，密码长度不足 40 时新增用户保存失败，满足长度时保存成功，且该用户可登录
# 测试步骤：
#   1. Minimum Password Length = 40，保存
#   2. 新增用户，Password 长度 < 40 → 保存失败
#   3. 新增用户，Password 长度 = 40 → 保存成功
#   4. 使用新用户登录
# 预期结果：
#   2. 保存失败
#   3. 保存成功
#   4. 登录成功
def test_TestCase_AcuHMI_007_03_case05_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    _create_role(page)
    try:
        _set_min_length(page, 40)

        assert not _try_create_user(page, _USER_NAME, _SHORT_PWD), \
            f"密码长度 {len(_SHORT_PWD)} < 40，新增用户应失败"

        assert _try_create_user(page, _USER_NAME, _GOOD_PWD), \
            f"密码长度 {len(_GOOD_PWD)} = 40，新增用户应成功"

        assert _can_login(browser, _USER_NAME, _GOOD_PWD), \
            f"用户 {_USER_NAME} 应能以 40 位密码登录"
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_length(page)
