import re
import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用户名/密码长度边界测试


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


def _delete_user_if_exists(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    # Use exact text match on username cell to avoid partial matches (e.g. "a" in "admin")
    row = page.locator("tbody").get_by_role("row").filter(
        has=page.get_by_text(username, exact=True)
    )
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _open_add_user(page):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)


def _cancel_add_user(page):
    try:
        page.get_by_role("button", name="Cancel").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


# 用例编号：TestCase_AcuHMI_007_01_case06
# 用例标题：添加用户，用户名长度为1，密码长度为5，保存配置失败；勾选Override password policy，添加成功
# 测试步骤：
#   1. Add User：Username="a"，Password="12345"（5位，不符合最短6位要求），Role=admin，保存
#   2. 预期失败
#   3. 勾选 "Override Password Policy"，保存
#   4. 预期成功
# 预期结果：
#   2. 保存失败，提示密码长度过短
#   4. 添加成功
def test_TestCase_AcuHMI_007_01_case06(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    short_user = "a"
    short_pwd = "12345"

    try:
        # Step 1: 短密码不勾选 Override → 应失败
        _open_add_user(page)
        page.get_by_label("Username", exact=True).fill(short_user)
        page.get_by_label("Password", exact=True).fill(short_pwd)
        page.get_by_label("Repeat Password", exact=True).fill(short_pwd)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name="admin").click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)

        # 对话框应保持打开（保存失败）
        assert page.get_by_label("Password", exact=True).is_visible(), \
            "密码长度5（低于最小6位）且未勾选Override时，添加应失败（对话框应保持）"

        # Step 2: 勾选 Override Password Policy → 应成功
        # ElementUI hides the actual <input>; toggle via .el-checkbox__inner
        page.locator(".el-form-item").filter(has_text="Override Password Policy").locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        assert not page.get_by_label("Password", exact=True).is_visible(), \
            "勾选Override Policy后，短密码创建用户应成功（对话框应关闭）"
    finally:
        _delete_user_if_exists(page, short_user)
