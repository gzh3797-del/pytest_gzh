import pytest
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


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
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case06_1
# 用例标题：添加用户，用户名长度为2，密码长度为8（符合策略），可添加成功
# 测试步骤：
#   1. Add User：Username="aa"，Password="Ab@123456"（符合策略：8位，含大小写+数字+特殊），Role=admin，保存
#   2. 预期成功
#   3. 重新创建同名用户，密码不符合策略（如"12345678"），不勾选Override，应失败
#   4. 勾选Override，应成功
# 预期结果：
#   2. 添加成功，可登录
def test_TestCase_AcuHMI_007_01_case06_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    username = "aa"
    good_pwd = "Ab@123456"

    try:
        # Step 1: 符合策略的密码 → 应成功
        _nav_to_submenu(page, "User Configuration")
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)
        page.get_by_label("Username", exact=True).fill(username)
        page.get_by_label("Password", exact=True).fill(good_pwd)
        page.get_by_label("Repeat Password", exact=True).fill(good_pwd)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name="admin").click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        assert not page.get_by_label("Password", exact=True).is_visible(), \
            "用户名=2位，密码=8位符合策略，添加应成功"

        # 验证可登录
        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(username)
            p.get_by_role("textbox", name="Enter Password").fill(good_pwd)
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
                f"用户 {username} 应能登录，当前 URL: {p.url}"
        finally:
            ctx.close()
    finally:
        _delete_user_if_exists(page, username)
