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
    """在独立浏览器上下文中尝试登录，返回是否登录成功。"""
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


# 用例编号：TestCase_AcuHMI_007_01_case03_2
# 用例标题：删除全部用户，登录所有用户
# 测试步骤（更新）：
#   1. 创建 uc032_1、uc032_2（view 角色）
#   2. 删除全部非 admin 用户
#   3. 被删除的用户尝试登录 → 应失败
#   4. admin 用户无法被删除（行无删除按钮，或点击后 admin 仍存在）
# 预期结果（更新）：
#   1. 被删除用户均无法登录成功
#   2. 最后1个 admin 用户提示无法删除
def test_TestCase_AcuHMI_007_01_case03_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    pwd = "Abc@12345"
    test_users = ["uc032_1", "uc032_2"]

    try:
        for username in test_users:
            _create_user(page, username, pwd, role="view")

        # Delete all non-admin test users
        for username in test_users:
            _delete_user(page, username)

        # Verify non-admin users are removed from list
        _nav_to_submenu(page, "User Configuration")
        for username in test_users:
            row = page.locator("tbody").get_by_role("row").filter(has_text=username)
            assert row.count() == 0, f"用户 {username} 应已被删除"

        # Step 1: Verify deleted users cannot login (independent browser contexts)
        for username in test_users:
            logged_in = _login_verify(browser, username, pwd)
            assert not logged_in, f"用户 {username} 已被删除，不应能登录成功"

        # Step 2: Verify admin cannot be deleted
        # Admin session is still active on `page`; navigate directly
        _nav_to_submenu(page, "User Configuration")
        admin_row = page.locator("tbody").get_by_role("row").filter(
            has_text="admin"
        ).first

        btn_count = admin_row.get_by_role("button").count()

        if btn_count > 0:
            # Click the last button (potential delete) and verify admin survives
            admin_row.get_by_role("button").last.click()
            page.wait_for_timeout(800)

            confirm_btn = page.get_by_role("button", name="Yes, continue")
            if confirm_btn.is_visible():
                confirm_btn.click()
                page.wait_for_timeout(800)

            # Admin must still be present regardless
            _nav_to_submenu(page, "User Configuration")
            admin_still_present = page.locator("tbody").get_by_role("row").filter(
                has_text="admin"
            ).count() > 0
            assert admin_still_present, "admin 用户应无法被删除，仍应存在于列表中"
        else:
            # No buttons on admin row — protected at UI level
            admin_still_present = page.locator("tbody").get_by_role("row").filter(
                has_text="admin"
            ).count() > 0
            assert admin_still_present, "admin 用户应存在于列表中（无删除按钮）"

    finally:
        for username in test_users:
            _delete_user(page, username)
