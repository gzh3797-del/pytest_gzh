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


def _create_user(page, username: str, password: str, role: str = "admin"):
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


# 用例编号：TestCase_AcuHMI_007_01_case02
# 用例标题：添加1、2、10个用户，登录用户
# 测试步骤：
#   1. 添加 test1（admin），保存，验证添加成功
#   2. 再添加 test2（admin）、test3（admin）、test4（view）、test5（view），共5个用户
#   3. 依次使用各用户登录，验证登录成功
# 预期结果：
#   所有用户均可登录成功，权限正确
def test_TestCase_AcuHMI_007_01_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    users = [
        ("uc02test1", "Abc@12345", "admin"),
        ("uc02test2", "Abc@12345", "admin"),
        ("uc02test3", "Abc@12345", "admin"),
        ("uc02test4", "Abc@12345", "view"),
        ("uc02test5", "Abc@12345", "view"),
    ]

    try:
        for username, pwd, role in users:
            _create_user(page, username, pwd, role)

        for username, pwd, role in users:
            assert _login_verify(browser, username, pwd), \
                f"用户 {username}（{role} 角色）应能登录成功"
    finally:
        for username, pwd, role in users:
            _delete_user(page, username)
