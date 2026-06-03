import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME = "pp01_3role"
_USER_NAME = "pp01_3user"
_INIT_PWD  = "Admin@110001"
# 密码含字母但无数字，在 Numbers+Letters 策略下应拒绝
_BAD_PWD   = "####abcde"

_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]

_IDX_UPPER_LOWER    = 0
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
        page.get_by_role("option", name="none", exact=True).click()
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


# 用例编号：TestCase_AcuHMI_007_03_case01_3
# 用例标题：配置密码策略，仅勾选数字和字母，创建用户密码不包含数字，无法创建用户
# 测试步骤：
#   1. Password Policy 仅勾选 Numbers and Letters，保存
#   2. 创建用户，密码为 "####abcde"（只有字母和特殊字符，无数字）
# 预期结果：
#   1. 保存成功
#   2. 创建用户失败
def test_TestCase_AcuHMI_007_03_case01_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _create_role(page)
    try:
        # Step 1: 仅启用 Numbers+Letters 策略
        _set_policy(page, upper_lower=False, numbers_letters=True, special_chars=False)

        # Step 2: 创建用户，密码无数字 → 应失败
        _nav_to_submenu(page, "User Configuration")
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)
        page.get_by_label("Username", exact=True).fill(_USER_NAME)
        page.get_by_label("Password", exact=True).fill(_BAD_PWD)
        page.get_by_label("Repeat Password", exact=True).fill(_BAD_PWD)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name=_ROLE_NAME).click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        assert page.get_by_label("Password", exact=True).is_visible(), \
            "Numbers+Letters 策略下，无数字的密码创建用户应失败（对话框应保持打开）"

        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
            page.wait_for_timeout(500)
        except Exception:
            page.keyboard.press("Escape")
    finally:
        _delete_user_if_exists(page, _USER_NAME)
        _delete_role(page)
        _restore_default_policy(page)
