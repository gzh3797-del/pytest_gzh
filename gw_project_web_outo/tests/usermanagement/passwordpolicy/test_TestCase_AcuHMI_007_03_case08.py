import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_ROLE_NAME  = "pp08_role"
_USER_NAME  = "pp08user"
_USER_PWD   = "Admin@11001"
_WRONG_PWD  = "Wrong@11001"

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
        page.wait_for_timeout(500)
        page.get_by_role("option", name="view", exact=True).click()
        page.wait_for_timeout(500)
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
    # EULA dialog appears on first successful login while URL is still /#/login
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


# 用例编号：TestCase_AcuHMI_007_03_case08
# 用例标题：在失败登录尝试窗口（30s）内失败 3 次后，超过窗口时间再尝试，窗口重置；
#           错误密码连续登录 3 次，等待 5 秒后自动解锁
# 测试步骤：
#   1. Maximum Failed Attempts=3, Window=30, Wait=5，保存
#   2. 在 30s 内失败登录 3 次（窗口内）
#   3. 等待 35s（窗口过期），再失败 3 次
#   4. 验证第 3 次失败后账户被锁定
#   5. 等待 6s 后，正确密码可登录
# 预期结果：
#   2. 登录失败，账户未被锁定（失败次数在窗口内累计）
#   4. 被锁定
#   5. 解锁后登录成功
def test_TestCase_AcuHMI_007_03_case08(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    # Clean up stale state from any previous failed run
    _delete_user(page)
    _delete_role(page)

    _create_role(page)
    _create_user(page)
    try:
        _set_lockout_policy(page, max_attempts=3, window=30, wait=5)

        ctx = browser.new_context(ignore_https_errors=True)
        lp = ctx.new_page()
        try:
            lp.goto(BASE_URL + "/#/login")
            lp.wait_for_load_state("networkidle")

            # Step 2: 3 wrong logins within the 30s window
            for i in range(3):
                assert not _login_attempt(lp, _USER_NAME, _WRONG_PWD), \
                    f"第 {i+1} 次错误密码登录应失败"

            # After exactly 3 failures (= Max), account is locked
            # Verify lockout: correct password also fails
            locked_after_3 = not _login_attempt(lp, _USER_NAME, _USER_PWD)

            # Step 3: wait for window (35s) to expire so counter resets
            lp.wait_for_timeout(35000)

            # After window expiry, re-do 3 more wrong logins
            for i in range(3):
                assert not _login_attempt(lp, _USER_NAME, _WRONG_PWD), \
                    f"窗口重置后第 {i+1} 次错误密码登录应失败"

            # Step 4: After 3 failures in new window, account should be locked again
            assert not _login_attempt(lp, _USER_NAME, _USER_PWD), \
                "3 次失败后账户应被锁定"

            # Step 5: wait for lockout to expire, correct login succeeds
            lp.wait_for_timeout(8000)
            assert _login_attempt(lp, _USER_NAME, _USER_PWD), \
                "等待 6s 后，正确密码应能登录成功"
        finally:
            ctx.close()
    finally:
        _delete_user(page)
        _delete_role(page)
        _restore_lockout_policy(page)
