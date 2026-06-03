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


def _lock_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").first.click()  # Lock is the first button for non-admin rows
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


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


# 用例编号：TestCase_AcuHMI_007_01_case04_1
# 用例标题：锁全部-1个用户，登录所有用户
# 测试步骤：
#   1. 创建5个用户（uc041_1 ~ uc041_5）
#   2. 锁定其中4个（保留 uc041_5 不锁定）
#   3. 被锁的4个用户登录 → 应失败
#   4. 未锁定的 uc041_5 登录 → 应成功
# 预期结果：
#   仅未锁定的1个用户可以登录
def test_TestCase_AcuHMI_007_01_case04_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    pwd = "Abc@12345"
    to_lock = [f"uc041_{i}" for i in range(1, 5)]
    to_keep = ["uc041_5"]
    all_users = to_lock + to_keep

    try:
        for username in all_users:
            _create_user(page, username, pwd, role="view")

        # Lock all-1 users
        for username in to_lock:
            _lock_user(page, username)

        # Locked users should NOT login
        for username in to_lock:
            assert not _login_verify(browser, username, pwd), \
                f"已锁定用户 {username} 不应能登录"

        # Unlocked user should login successfully
        assert _login_verify(browser, to_keep[0], pwd), \
            f"未锁定用户 {to_keep[0]} 应能正常登录"
    finally:
        for username in all_users:
            _delete_user(page, username)
