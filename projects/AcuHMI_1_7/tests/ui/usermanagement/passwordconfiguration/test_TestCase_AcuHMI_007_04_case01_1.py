import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_TEST_USER = "pwdcfg01b"
_INIT_PWD  = "Admin@110001"


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


def _open_pm_edit_form(page, username: str):
    _nav_to_submenu(page, "Password Management")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").click()
    page.wait_for_timeout(500)


def _is_form_still_open(page) -> bool:
    return page.get_by_label("Password", exact=True).is_visible()


# ── 用例编号：TestCase_AcuHMI_007_04_case01_1
# 用例标题：修改密码，密码长度为5、129，保存配置失败，提示错误信息正确
# 预置条件：管理权限登录 AcuHMI 网页
# 测试步骤：
#   1. Password Management 页面选择任一用户，点击编辑按钮
#   2. 输入长度为 5 的密码（12345），点击 Save
#   3. 验证保存失败（表单不关闭 / 出现错误提示）
#   4. 清空密码框，输入长度为 129 的密码，点击 Save
#   5. 验证保存失败
# 预期结果：
#   2. 保存配置失败
#   4. 保存配置失败
def test_TestCase_AcuHMI_007_04_case01_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _create_user(page, _TEST_USER, _INIT_PWD)
    try:
        # ── 场景 1：密码长度 5（低于最小长度 6）────────────────────────
        _open_pm_edit_form(page, _TEST_USER)
        pwd_5 = "12345"
        page.get_by_label("Password", exact=True).fill(pwd_5)
        page.get_by_label("Repeat Password", exact=True).fill(pwd_5)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)
        assert _is_form_still_open(page), \
            "密码长度 5 应保存失败，但表单已关闭（保存成功），用例失败"

        # 关闭表单，准备下一场景
        try:
            page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            page.keyboard.press("Escape")
        page.wait_for_timeout(500)

        # ── 场景 2：密码长度 129（超过最大长度 128）─────────────────────
        _open_pm_edit_form(page, _TEST_USER)
        pwd_129 = "Admin@12" + "x" * 121   # 8 + 121 = 129 位
        page.get_by_label("Password", exact=True).fill(pwd_129)
        page.get_by_label("Repeat Password", exact=True).fill(pwd_129)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)
        assert _is_form_still_open(page), \
            "密码长度 129 应保存失败，但表单已关闭（保存成功），用例失败"

        # 关闭表单
        try:
            page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            page.keyboard.press("Escape")
    finally:
        _delete_user(page, _TEST_USER)