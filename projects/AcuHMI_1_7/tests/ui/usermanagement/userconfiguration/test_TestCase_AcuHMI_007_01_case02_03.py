import pytest
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


def _count_users(page) -> int:
    _nav_to_submenu(page, "User Configuration")
    return page.locator("tbody").get_by_role("row").count()


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
    try:
        page.get_by_label("Password", exact=True).wait_for(state="hidden", timeout=5000)
    except Exception:
        pass


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_01_case02_03
# 用例标题：添加最多限制个用户（32个），验证登录
# 测试步骤：
#   系统最大用户数为32（含 admin）；动态计算当前剩余名额并填满
#   1. 统计当前已有用户数 current_count
#   2. 循环创建 (32 - current_count) 个用户，直到达到上限
#   3. 验证每个用户均可成功添加（对话框关闭 = 成功）
#   4. 验证总用户数达到 32
# 预期结果：
#   创建的所有用户均添加成功，总用户数达到 32
def test_TestCase_AcuHMI_007_01_case02_03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    MAX_TOTAL = 32
    pwd = "Abc@12345"
    created = []

    try:
        # 统计当前用户数，动态计算可新增名额
        current_count = _count_users(page)
        slots_available = MAX_TOTAL - current_count
        assert slots_available > 0, \
            f"当前用户数 {current_count} 已达或超过最大限制 {MAX_TOTAL}，无法测试添加上限"

        for i in range(1, slots_available + 1):
            username = f"uc203_{i:02d}"
            _create_user(page, username, pwd, role="view")
            dialog_still_open = page.get_by_label("Password", exact=True).is_visible()
            assert not dialog_still_open, \
                f"第 {i} 个用户 {username} 应添加成功（当前共 {current_count + i} 个），但对话框仍开启"
            created.append(username)

        # 验证总用户数达到 32
        final_count = _count_users(page)
        assert final_count == MAX_TOTAL, \
            f"添加后总用户数应为 {MAX_TOTAL}，实际为 {final_count}"
    finally:
        for username in created:
            _delete_user(page, username)
