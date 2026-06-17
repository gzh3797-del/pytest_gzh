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


def _create_user_with_options(page, username: str, password: str, role: str,
                               multiple_login: bool = False,
                               override_policy: bool = False):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role).click()
    page.wait_for_timeout(300)
    # "Multiple Login" 默认 = True(ON)；"Override Password Policy" 默认 = False(OFF)
    # 仅在目标值与默认值不同时才点击以切换
    if not multiple_login:   # 需要 OFF，从默认 ON 切换
        page.locator(".el-form-item").filter(has_text="Multiple Login").locator(".el-checkbox__inner").click()
        page.wait_for_timeout(200)
    if override_policy:      # 需要 ON，从默认 OFF 切换
        page.locator(".el-form-item").filter(has_text="Override Password Policy").locator(".el-checkbox__inner").click()
        page.wait_for_timeout(200)
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


def _login_in_ctx(browser, username: str, password: str):
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
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
    return ctx, p


# 用例编号：TestCase_AcuHMI_007_01_case02_01
# 用例标题：添加用户，验证多重登录和覆盖密码策略选项
# 测试步骤：
#   1. 添加 user1（admin，勾选 Multiple Login + Override Password Policy）
#   2. 用两个浏览器窗口同时登录 user1，均应成功
#   3. 添加 user2（admin，不勾选 Multiple Login，不勾选 Override Password Policy）
#   4. 用两个浏览器窗口登录 user2：第一个成功，第二个应失败或提示已登录
# 预期结果：
#   user1 多窗口均登录成功；user2 仅第一个窗口登录成功
def test_TestCase_AcuHMI_007_01_case02_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    user1 = "uc0201multi"
    user2 = "uc0201single"
    pwd = "Abc@12345"

    try:
        # user1: Multiple Login + Override Password Policy
        _create_user_with_options(page, user1, pwd, "admin",
                                   multiple_login=True, override_policy=True)

        ctx1a, p1a = _login_in_ctx(browser, user1, pwd)
        ctx1b, p1b = _login_in_ctx(browser, user1, pwd)
        try:
            assert "/#/login" not in p1a.url, \
                f"user1（Multiple Login=on）第一窗口应登录成功，URL: {p1a.url}"
            assert "/#/login" not in p1b.url, \
                f"user1（Multiple Login=on）第二窗口也应登录成功，URL: {p1b.url}"
        finally:
            ctx1a.close()
            ctx1b.close()

        # user2: Multiple Login off, Override Password Policy off
        _create_user_with_options(page, user2, pwd, "admin",
                                   multiple_login=False, override_policy=False)

        ctx2a, p2a = _login_in_ctx(browser, user2, pwd)
        try:
            assert "/#/login" not in p2a.url, \
                f"user2 第一个窗口应登录成功，URL: {p2a.url}"

            ctx2b, p2b = _login_in_ctx(browser, user2, pwd)
            try:
                # 第二个窗口：不允许多重登录时，应登录失败或已有会话被踢出
                second_login_ok = "/#/login" not in p2b.url
                if second_login_ok:
                    # 若第二次登录成功，第一个会话应被踢出
                    # 等待 dashboard 轮询检测到 token 失效（最多 25s）
                    p2a.wait_for_timeout(25000)
                    # SPA 踢出可能以 Warning 对话框或跳转到登录页两种方式显示
                    first_kicked = (
                        "login" in p2a.url.lower()
                        or p2a.locator(".el-overlay-message-box").count() > 0
                    )
                    assert first_kicked, \
                        f"user2 不允许多重登录：第二次登录成功时，第一个会话应被踢出（URL={p2a.url}）"
                else:
                    # 第二次登录直接失败也是合法行为
                    pass
            finally:
                ctx2b.close()
        finally:
            ctx2a.close()
    finally:
        _delete_user(page, user1)
        _delete_user(page, user2)
