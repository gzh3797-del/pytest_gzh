import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_USER_EDIT = "eula06edit"
_USER_VIEW = "eula06view"
_INIT_PWD  = "Admin@110001"

_EULA_ACCEPT_BTN  = "Accept"
_EULA_CLOSE_BTN   = "Close"   # EULA 不接受按钮，点击后停留在登录页


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
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _verify_eula_flow(browser, username: str, password: str):
    """
    验证 EULA 完整流程：
    1. 首次登录 → EULA 弹框出现，点击 Close（不接受）→ 停留在登录页
    2. 再次登录 → 点击 Accept（接受）→ 进入系统成功
    """
    # ── 第一次登录：点击 Close（不接受 EULA）────────────────────────────────
    ctx1 = browser.new_context(ignore_https_errors=True)
    p1   = ctx1.new_page()
    try:
        p1.goto(BASE_URL + "/#/login")
        p1.wait_for_load_state("networkidle")
        p1.get_by_role("textbox", name="Enter User Name").fill(username)
        p1.get_by_role("textbox", name="Enter Password").fill(password)
        p1.get_by_role("button", name="Sign In").click()
        p1.wait_for_load_state("networkidle")
        p1.wait_for_timeout(1000)

        # EULA 弹框应出现（有 Accept 和 Close 按钮）
        assert p1.get_by_role("button", name=_EULA_ACCEPT_BTN).is_visible(timeout=5000), \
            f"用户 {username} 首次登录后 EULA 弹框（Accept 按钮）未出现"

        # 点击 Close（不接受）
        p1.get_by_role("button", name=_EULA_CLOSE_BTN).click()
        p1.wait_for_load_state("networkidle")
        p1.wait_for_timeout(500)

        # Close 后应停留在登录页，无法进入系统
        assert "/#/login" in p1.url, \
            f"用户 {username} 点击 EULA Close 后应停留在登录页，当前 URL: {p1.url}"
    finally:
        ctx1.close()

    # ── 第二次登录：点击 Accept（接受 EULA）────────────────────────────────
    ctx2 = browser.new_context(ignore_https_errors=True)
    p2   = ctx2.new_page()
    try:
        p2.goto(BASE_URL + "/#/login")
        p2.wait_for_load_state("networkidle")
        p2.get_by_role("textbox", name="Enter User Name").fill(username)
        p2.get_by_role("textbox", name="Enter Password").fill(password)
        p2.get_by_role("button", name="Sign In").click()
        p2.wait_for_load_state("networkidle")
        p2.wait_for_timeout(1000)

        # 接受 EULA
        assert p2.get_by_role("button", name=_EULA_ACCEPT_BTN).is_visible(timeout=5000), \
            f"用户 {username} 第二次登录 EULA 未出现"
        p2.get_by_role("button", name=_EULA_ACCEPT_BTN).click()
        p2.wait_for_load_state("networkidle")
        p2.wait_for_timeout(500)

        # 处理默认密码提示
        try:
            p2.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass

        assert "/#/login" not in p2.url, \
            f"用户 {username} 接受 EULA 后应进入系统，当前 URL: {p2.url}"
    finally:
        ctx2.close()


# 用例编号：TestCase_AcuHMI_007_01_case06_6
# 用例标题：验证新建用户登录系统，是否需要接受EULA
# 预置条件：服务启动正常，admin 账号登录成功，进入用户设置界面
# 测试步骤：
#   1. User Configuration 新建 edit 权限用户 eula06edit
#   2. eula06edit 登录 → 出现 EULA 弹框 → 点击拒绝 → 无法登录
#   3. 重新登录 → 点击接受 → 登录成功
#   4. User Configuration 新建 view 权限用户 eula06view
#   5. eula06view 登录 → 出现 EULA 弹框 → 点击拒绝 → 无法登录
#   6. 重新登录 → 点击接受 → 登录成功
# 预期结果：
#   拒绝 EULA → 无法登录；接受 EULA → 登录成功
def test_TestCase_AcuHMI_007_01_case06_6(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser    = admin_page.context.browser

    _create_user(admin_page, _USER_EDIT, _INIT_PWD, role="view")
    _create_user(admin_page, _USER_VIEW, _INIT_PWD, role="view")
    try:
        _verify_eula_flow(browser, _USER_EDIT, _INIT_PWD)
        _verify_eula_flow(browser, _USER_VIEW, _INIT_PWD)
    finally:
        _delete_user(admin_page, _USER_EDIT)
        _delete_user(admin_page, _USER_VIEW)
