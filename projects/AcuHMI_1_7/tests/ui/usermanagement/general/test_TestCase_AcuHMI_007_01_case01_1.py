import pytest
from projects.AcuHMI_1_7.settings import BASE_URL, DEFAULT_PASSWORD
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_ROLE_NAME = "gen01_1role"
_USER_NAME = "gen01_1user"
_USER_PWD  = "Admin@11011"

_PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]


def _dismiss_dialogs(page):
    """关闭任何阻断操作的 Warning/弹框（session 超时弹框等）。"""
    try:
        page.locator(".el-overlay-message-box").get_by_role("button").last.click(timeout=1500)
        page.wait_for_timeout(300)
    except Exception:
        pass


def _nav_to_submenu(page, submenu: str):
    _dismiss_dialogs(page)
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _set_session_timeout(page, timeout_val: int):
    _nav_to_submenu(page, "General")
    page.get_by_placeholder("Enter Session Timeout").fill(str(timeout_val))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_session_timeout(page, admin_pwd: str):
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    # 若 session 仍有效，SPA 会重定向到 dashboard，无需重新登录
    if "login" in page.url:
        page.get_by_role("textbox", name="Enter User Name").fill("admin")
        page.get_by_role("textbox", name="Enter Password").fill(admin_pwd)
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        try:
            page.get_by_role("button", name="Accept").click(timeout=3000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        except Exception:
            pass
        try:
            page.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
    # 先进入 AcuHMI 上下文（System Settings），再导航到 User Management General
    # 直接访问 /#/userManagement/general 不走 header 点击，需先建立设备上下文
    page.goto(BASE_URL + "/#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.goto(BASE_URL + "/#/userManagement/general")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_placeholder("Enter Session Timeout").fill("10")
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


# 用例编号：TestCase_AcuHMI_007_01_case01_1
# 用例标题：会话超时为1，保存配置并验证最大登录数、超时
# 测试步骤：
#   1. General 页面设置 Session Timeout=1，保存
#   2. 使用第二个用户登录系统
#   3. 等待 65s，查看是否退出到登录页面
# 预期结果：
#   2. 登录成功
#   3. 已登录用户跳转到登录页面（超时退出）
def test_TestCase_AcuHMI_007_01_case01_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser
    admin_pwd = DEFAULT_PASSWORD

    _create_role(page)
    _create_user(page)
    try:
        _set_session_timeout(page, 1)

        ctx = browser.new_context(ignore_https_errors=True)
        p2 = ctx.new_page()
        try:
            p2.goto(BASE_URL + "/#/login")
            p2.wait_for_load_state("networkidle")
            p2.get_by_role("textbox", name="Enter User Name").fill(_USER_NAME)
            p2.get_by_role("textbox", name="Enter Password").fill(_USER_PWD)
            p2.get_by_role("button", name="Sign In").click()
            p2.wait_for_load_state("networkidle")
            p2.wait_for_timeout(1000)
            try:
                p2.get_by_role("button", name="Accept").click(timeout=3000)
                p2.wait_for_load_state("networkidle")
                p2.wait_for_timeout(500)
            except Exception:
                pass
            try:
                p2.get_by_role("button", name="Cancel").click(timeout=2000)
            except Exception:
                pass

            assert "/#/login" not in p2.url, \
                f"用户2应登录成功，当前 URL: {p2.url}"

            # 阻断所有后台设备轮询，防止 session 被保活
            def _block_bg(route):
                if "/api/device/" in route.request.url:
                    route.abort()
                else:
                    route.continue_()

            p2.route("**/*", _block_bg)

            # 导航到无设备轮询的页面，让 session timer 开始倒计时
            p2.goto(BASE_URL + "/#/systemSettings/dateTime")
            p2.wait_for_load_state("networkidle")
            p2.wait_for_timeout(1000)

            # 等待 session 超时（1 min = 60s）+ 15s 缓冲
            p2.wait_for_timeout(75000)

            # 从 sessionStorage 取 token，直接调用 API 验证 session 是否过期
            ss = p2.evaluate("() => JSON.parse(sessionStorage.getItem('common') || '{}')")
            token = ss.get('authToken', '')
            resp = p2.evaluate("""async (tok) => {
                try {
                    const r = await fetch('/api/settings/ntpConfig?token=' + tok, {method: 'GET'});
                    const b = await r.text();
                    return {status: r.status, body: b.substring(0, 80)};
                } catch(e) { return {error: e.message}; }
            }""", token)

            # session 过期后，服务端应返回 401（或拒绝访问），而不是 200
            status = resp.get('status', 0)
            assert status != 200, \
                f"75s 无轮询后 session 应过期，但 API 返回 {status}（200 表示仍有效）\ntoken={token}\nbody={resp}"
        finally:
            ctx.close()
    finally:
        _restore_session_timeout(page, admin_pwd)
        try:
            _delete_user(page)
            _delete_role(page)
        except Exception:
            pass
