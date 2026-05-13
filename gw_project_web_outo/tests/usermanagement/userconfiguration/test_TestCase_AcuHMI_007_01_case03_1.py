import pytest
from config.settings import BASE_URL
from pages.login_page import LoginPage


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
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_verify(browser, username: str, password: str) -> bool:
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


# 用例编号：TestCase_AcuHMI_007_01_case03_1
# 用例标题：删除全部-1个用户，登录所有用户
# 测试步骤：
#   1. 创建 5 个用户（uc031_1 ~ uc031_5），共 5 个新用户（admin 已存在）
#   2. 删除其中 4 个用户（保留 uc031_5）
#   3. 被删除的 4 个用户登录 → 应失败
#   4. 保留的 uc031_5 登录 → 应成功
# 预期结果：
#   仅保留的1个用户可以登录，其余已删除用户均无法登录
def test_TestCase_AcuHMI_007_01_case03_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    pwd = "Abc@12345"
    to_delete = [f"uc031_{i}" for i in range(1, 5)]  # 4 users to delete
    to_keep = ["uc031_5"]
    all_users = to_delete + to_keep

    try:
        for username in all_users:
            _create_user(page, username, pwd, role="view")

        # Delete all-1 users (keep the last one)
        for username in to_delete:
            _delete_user(page, username)

        # Deleted users should NOT login
        for username in to_delete:
            assert not _login_verify(browser, username, pwd), \
                f"已删除用户 {username} 不应能登录"

        # Kept user should login
        assert _login_verify(browser, to_keep[0], pwd), \
            f"保留的用户 {to_keep[0]} 应能正常登录"
    finally:
        for username in to_keep:
            _delete_user(page, username)
