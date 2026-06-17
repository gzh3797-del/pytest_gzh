import pytest
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "pp0901role"
_USER_NAME = "pp0901user"
_USER_PWD  = "Admin@11091"
_WRONG_PWD = "Wrong@11091"

_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]


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


def _set_lockout_policy(page, max_attempts, window, wait):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Maximum Failed Attempts").fill(str(max_attempts))
    page.get_by_placeholder("Enter Failed Login Attempt Window").fill(str(window))
    page.get_by_placeholder("Enter Failed Login Wait").fill(str(wait))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_lockout_policy(page):
    _nav_to_submenu(page, "Password Policy")
    page.get_by_placeholder("Enter Maximum Failed Attempts").fill("0")
    page.get_by_placeholder("Enter Failed Login Attempt Window").fill("0")
    page.get_by_placeholder("Enter Failed Login Wait").fill("0")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _create_role(page):
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_ROLE_NAME)
    for lbl in _PERM_LABELS:
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="view", exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=_ROLE_NAME)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.wait_for_timeout(500)
    try:
        page.get_by_role("button", name="Yes, continue").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


def _create_user(page):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(_USER_NAME)
    page.get_by_label("Password", exact=True).fill(_USER_PWD)
    page.get_by_label("Repeat Password", exact=True).fill(_USER_PWD)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=_ROLE_NAME).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user(page):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=_USER_NAME)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_attempt(page, username, password) -> bool:
    page.get_by_role("textbox", name="Enter User Name").fill(username)
    page.get_by_role("textbox", name="Enter Password").fill(password)
    page.get_by_role("button", name="Sign In").click(force=True)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if page.locator(".el-dialog").filter(has_text="LICENSE AGREEMENT").count() > 0:
        page.get_by_role("button", name="Accept").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    if "/#/login" not in page.url:
        try:
            page.get_by_role("button", name="Cancel").click(timeout=1000)
        except Exception:
            pass
    return "/#/login" not in page.url


# 用例编号：TestCase_AcuHMI_007_03_case09_01
# 用例标题：用户在失败登录尝试窗口（10分钟）内连续输入错误密码5次（超过阈值），验证账户是否被锁定
# 测试步骤：
#   1. Failed Login Wait=30，Failed Login Attempt Window=600，Maximum Failed Attempts=5，保存
#   2. 使用错误密码登录 5 次
#   3. 尝试第 6 次登录（用正确密码）
#   4. 等待 31s（Wait=30 过期），再次用正确密码登录
# 预期结果：
#   1. 设置成功
#   3. 第 6 次登录被锁定拒绝
#   4. 等待 31s 后解锁，登录成功
def test_TestCase_AcuHMI_007_03_case09_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    _delete_user(page)
    _delete_role(page)
    _create_role(page)
    _create_user(page)
    try:
        _set_lockout_policy(page, max_attempts=5, window=600, wait=30)

        ctx = browser.new_context(ignore_https_errors=True)
        lp = ctx.new_page()
        try:
            lp.goto(BASE_URL + "/#/login")
            lp.wait_for_load_state("networkidle")

            # 连续失败 5 次
            for i in range(5):
                assert not _login_attempt(lp, _USER_NAME, _WRONG_PWD), \
                    f"第 {i+1} 次错误密码登录应失败"

            # 第 6 次（正确密码）应被锁定
            assert not _login_attempt(lp, _USER_NAME, _USER_PWD), \
                "5 次失败后账户应被锁定，正确密码也应被拒绝"

            # 等待 Wait=30s 过期后解锁
            lp.wait_for_timeout(32000)
            assert _login_attempt(lp, _USER_NAME, _USER_PWD), \
                "等待 32s 后账户应解锁，正确密码应能登录成功"
        finally:
            ctx.close()
    finally:
        _delete_user(page)
        _delete_role(page)
        _restore_lockout_policy(page)
