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
    # 用户名可能被截断为40字符，用前20字符过滤
    row = page.locator("tbody").get_by_role("row").filter(has_text=username[:20])
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case06_2
# 用例标题：添加用户，用户名长度为40，密码长度为64，可添加成功并登录
# 测试步骤：
#   1. Add User：Username=40字符，Password=64字符（含大小写+数字+特殊），保存
#   2. 预期成功，可登录
# 预期结果：
#   登录成功
def test_TestCase_AcuHMI_007_01_case06_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    # 40字符用户名
    username = "qwertyuiopasdfghjklzxQWE01234567890123"  # 38 chars, pad to 40
    username = username[:38] + "ab"  # exactly 40
    # 64字符密码（含大小写+数字+特殊）
    pwd = "Abc@1234" + "qwertyuiopasdfghjklzxcvbnm0123456789012345678901234567"  # 8+54=62
    pwd = pwd[:60] + "Ab@1"  # 64 chars total

    try:
        _nav_to_submenu(page, "User Configuration")
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)
        page.get_by_label("Username", exact=True).fill(username)
        page.get_by_label("Password", exact=True).fill(pwd)
        page.get_by_label("Repeat Password", exact=True).fill(pwd)
        page.get_by_text("--Select Role--", exact=True).click()
        page.get_by_role("option", name="admin").click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1500)

        assert not page.get_by_label("Password", exact=True).is_visible(), \
            "用户名=40字符，密码=64字符，应添加成功"

        # 验证可登录（用实际输入后的用户名，可能被截断）
        ctx = browser.new_context(ignore_https_errors=True)
        p = ctx.new_page()
        try:
            p.goto(BASE_URL + "/#/login")
            p.wait_for_load_state("networkidle")
            p.get_by_role("textbox", name="Enter User Name").fill(username)
            p.get_by_role("textbox", name="Enter Password").fill(pwd)
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
                f"用户名40字符+密码64字符的用户应能登录，当前 URL: {p.url}"
        finally:
            ctx.close()
    finally:
        _delete_user_if_exists(page, username)
