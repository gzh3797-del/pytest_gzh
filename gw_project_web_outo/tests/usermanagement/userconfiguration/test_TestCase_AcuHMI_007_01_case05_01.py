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


def _set_checkbox(page, label: str, want: bool):
    cb = page.locator(".el-form-item").filter(has_text=label).locator(".el-checkbox")
    is_checked = "is-checked" in (cb.get_attribute("class") or "")
    if want != is_checked:
        cb.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)


def _fill_if_visible(page, placeholder: str, value: str):
    try:
        inp = page.get_by_placeholder(placeholder)
        if inp.is_visible(timeout=1000):
            inp.fill(value)
    except Exception:
        pass


def _create_user(page, username: str, password: str, role: str = "admin",
                  multiple_login: bool = False,
                  override_policy: bool = False,
                  override_expire: bool = False,
                  pw_expire: str = "0",
                  grace_period: str = "0",
                  min_age: str = "0"):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role).click()
    page.wait_for_timeout(300)

    _set_checkbox(page, "Multiple Login", multiple_login)
    _set_checkbox(page, "Override Password Policy", override_policy)
    _set_checkbox(page, "Override Password Expire", override_expire)

    if override_expire:
        page.wait_for_timeout(300)
        _fill_if_visible(page, "Enter Password Expires", pw_expire)
        _fill_if_visible(page, "Enter Grace Period", grace_period)
        _fill_if_visible(page, "Enter Minimum Password Age", min_age)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _edit_user(page, username: str,
               multiple_login: bool = False,
               override_policy: bool = False,
               override_expire: bool = False,
               pw_expire: str = "0",
               grace_period: str = "0",
               min_age: str = "0"):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    row.get_by_role("button").nth(1).click()
    page.wait_for_timeout(1000)

    # Set numeric fields BEFORE unchecking Override Password Expire
    # (unchecking hides them)
    if not override_expire:
        _fill_if_visible(page, "Enter Password Expires", pw_expire)
        _fill_if_visible(page, "Enter Grace Period", grace_period)
        _fill_if_visible(page, "Enter Minimum Password Age", min_age)

    _set_checkbox(page, "Override Password Expire", override_expire)

    if override_expire:
        page.wait_for_timeout(300)
        _fill_if_visible(page, "Enter Password Expires", pw_expire)
        _fill_if_visible(page, "Enter Grace Period", grace_period)
        _fill_if_visible(page, "Enter Minimum Password Age", min_age)

    _set_checkbox(page, "Override Password Policy", override_policy)
    _set_checkbox(page, "Multiple Login", multiple_login)

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


def _do_login(p, username: str, password: str):
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


def _login_in_ctx(browser, username: str, password: str):
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    _do_login(p, username, password)
    return ctx, p


# 用例编号：TestCase_AcuHMI_007_01_case05_01
# 用例标题：多重登录验证 — Multiple Login 开启/关闭
# 测试步骤：
#   1. 创建 user4（admin，勾选 Multiple Login + Override Password Policy + Override Password Expire，
#      Password Expires=1，Grace Period=1，Minimum Password Age=1）
#   2. 使用多个窗口登录 user4 → 均成功
#   3. 编辑 user4，取消勾选 Multiple Login + Override Password Policy + Override Password Expire，
#      Password Expires=0，Grace Period=0，Minimum Password Age=0
#   4. 浏览器登录 user4（Tab 1），同一浏览器新开页签再次登录（Tab 2，共享 Session）
#      → Tab 1 弹框提示 "Unauthenticated user, please log in!"
# 预期结果：
#   2. Multiple Login=on 时多窗口均成功
#   4. Tab 2 登录后，Tab 1 收到 Unauthenticated 弹框（同一浏览器 Session 被更新）
def test_TestCase_AcuHMI_007_01_case05_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    username = "uc0501user"
    pwd = "Abc@12345"

    try:
        # Step 1: 创建用户，勾选 Multiple Login + Override Policy + Override Expire
        _create_user(page, username, pwd, role="admin",
                     multiple_login=True,
                     override_policy=True,
                     override_expire=True,
                     pw_expire="1", grace_period="1", min_age="1")

        # Step 2: 多窗口登录 → 均成功（Multiple Login=on）
        ctx1, p1 = _login_in_ctx(browser, username, pwd)
        ctx2, p2 = _login_in_ctx(browser, username, pwd)
        try:
            assert "/#/login" not in p1.url, \
                "Multiple Login=on，第1个窗口应登录成功"
            assert "/#/login" not in p2.url, \
                "Multiple Login=on，第2个窗口应登录成功"
        finally:
            ctx1.close()
            ctx2.close()

        # Step 3: 编辑用户，取消 Multiple Login + Override Policy + Override Expire，
        #         Password Expires=0，Grace Period=0，Min Age=0
        _edit_user(page, username,
                   multiple_login=False,
                   override_policy=False,
                   override_expire=False,
                   pw_expire="0", grace_period="0", min_age="0")

        # Step 4: 同一浏览器 Context，Tab 1 先登录，再新开 Tab 2 登录
        #         Tab 2 登录后 Session Token 更新，Tab 1 的 Token 失效
        #         → Tab 1 应弹出 "Unauthenticated user, please log in!"
        ctx3 = browser.new_context(ignore_https_errors=True)
        try:
            p3 = ctx3.new_page()
            _do_login(p3, username, pwd)
            assert "/#/login" not in p3.url, \
                "Multiple Login=off，Tab 1 应先登录成功"

            # 同一 Context 新开 Tab 2 并登录 → 共享 Cookie，覆盖 Session
            p4 = ctx3.new_page()
            _do_login(p4, username, pwd)
            assert "/#/login" not in p4.url, \
                "Tab 2 在同一浏览器中应也能登录成功（触发 Session 刷新）"

            # 等待 Tab 1 Session 失效，检查弹框 / Toast
            p3.wait_for_timeout(3000)
            try:
                p3.reload()
                p3.wait_for_load_state("networkidle")
                p3.wait_for_timeout(1000)
            except Exception:
                pass

            # 检查 "Unauthenticated" 弹框（ElementUI MessageBox）或 Toast
            unauthenticated = p3.get_by_text("Unauthenticated user, please log in!",
                                              exact=False)
            try:
                unauthenticated.wait_for(state="visible", timeout=5000)
                assert unauthenticated.is_visible(), \
                    "Tab 1 应弹出 'Unauthenticated user, please log in!' 提示"
            except Exception:
                # 如果没有弹框，检查是否跳转到登录页
                assert "/#/login" in p3.url, \
                    "Tab 2 登录后，Tab 1 应弹出 Unauthenticated 提示或跳转到登录页"
        finally:
            ctx3.close()
    finally:
        _delete_user(page, username)
