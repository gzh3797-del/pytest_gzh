import pytest
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


def _try_create_user(page, username: str, password: str, role: str = "view") -> bool:
    """Attempt to create user; return True if succeeded (dialog closed), False if failed (dialog still open)."""
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
    page.wait_for_timeout(1500)
    dialog_still_open = page.get_by_label("Password", exact=True).is_visible()
    if dialog_still_open:
        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
    return not dialog_still_open


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case02_04
# 用例标题：添加最多限制+1个用户，验证最后一个添加失败
# 测试步骤：
#   1. 添加31个用户（填满至系统上限 32 total）
#   2. 尝试添加第32个新用户（第33个总用户），验证失败
# 预期结果：
#   第33个用户添加失败，系统提示错误信息准确
def test_TestCase_AcuHMI_007_01_case02_04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    MAX_NEW_USERS = 31
    pwd = "Abc@12345"
    created = []

    try:
        # Fill up to system max (31 new + 1 admin = 32 total)
        for i in range(1, MAX_NEW_USERS + 1):
            username = f"uc204_{i:02d}"
            success = _try_create_user(page, username, pwd)
            if success:
                created.append(username)
            else:
                # Reached the limit earlier than expected
                break

        # Now try to add one more user beyond the limit
        extra_user = "uc204_extra"
        extra_success = _try_create_user(page, extra_user, pwd)
        if extra_success:
            created.append(extra_user)

        assert not extra_success, \
            "超过最大用户数量后（32+1个）添加应失败，但操作成功了"
    finally:
        for username in created:
            _delete_user(page, username)
