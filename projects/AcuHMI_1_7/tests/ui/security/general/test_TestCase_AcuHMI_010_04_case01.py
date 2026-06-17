import pytest
from projects.AcuHMI_1_7.settings import BASE_URL, DEFAULT_PASSWORD, VIEW_USERNAME, VIEW_PASSWORD
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_VIEW_USERNAME = VIEW_USERNAME
_VIEW_PASSWORD = VIEW_PASSWORD


def _dismiss_dialogs(page):
    """关闭任何阻断操作的弹框（密码提示、EULA 等）。"""
    try:
        page.locator(".el-overlay-message-box").get_by_role("button").last.click(timeout=1500)
        page.wait_for_timeout(300)
    except Exception:
        pass


def _nav_to_user_management_general(page):
    _dismiss_dialogs(page)
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="General").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _set_session_timeout(page, minutes: int):
    _nav_to_user_management_general(page)
    page.get_by_placeholder("Enter Session Timeout").fill(str(minutes))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _restore_session_timeout(page):
    """重新登录 admin（若 session 已过期）并恢复 Session Timeout 为 10 分钟"""
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    if "login" in page.url:
        page.get_by_role("textbox", name="Enter User Name").fill("admin")
        page.get_by_role("textbox", name="Enter User Name").press("Tab")
        page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        _dismiss_dialogs(page)
    page.goto(BASE_URL + "/#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.goto(BASE_URL + "/#/userManagement/general")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_placeholder("Enter Session Timeout").fill("10")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _login_as_view(browser):
    """在独立 browser context 中以 view 用户登录，返回 (ctx, page)"""
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    p.goto(BASE_URL + "/#/login")
    p.wait_for_load_state("networkidle")
    p.get_by_role("textbox", name="Enter User Name").fill(_VIEW_USERNAME)
    p.get_by_role("textbox", name="Enter User Name").press("Tab")
    p.get_by_role("textbox", name="Enter Password").fill(_VIEW_PASSWORD)
    p.get_by_role("button", name="Sign In").click()
    p.wait_for_load_state("networkidle")
    p.wait_for_timeout(1500)
    # 处理 EULA 弹窗
    try:
        p.get_by_role("button", name="Accept", exact=True).click(timeout=3000)
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(1000)
    except Exception:
        pass
    # 关闭"使用默认密码"弹窗
    try:
        p.get_by_role("button", name="Close", exact=True).click(timeout=3000)
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(500)
    except Exception:
        pass
    return ctx, p


# 用例编号：TestCase_AcuHMI_010_04_case01
# 用例标题：登录成功后，长时间未操作，自动化注销会话并自动退出登录
# 预置条件：1.服务启动正常，账号登录成功
# 测试步骤：
#   1.1 admin 用户登录，设置 Session Timeout = 1 分钟
#   2.  view 用户在独立 context 中登录（验证新建 session 受超时约束）
#   3.  阻断设备轮询，等待 1.1 分钟（75s）
# 预期结果：
#   1. 长时间未操作自动注销会话
#   2. 等待 75s 后 view 用户的 session token 应返回 401
def test_TestCase_AcuHMI_010_04_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _dismiss_dialogs(page)

    assert "/#/login" not in page.url, f"admin 应登录成功，当前URL: {page.url}"

    # admin 设置 Session Timeout = 1 分钟
    _set_session_timeout(page, 1)

    browser = page.context.browser
    ctx, view_page = _login_as_view(browser)
    try:
        assert "/#/login" not in view_page.url, \
            f"view 用户应登录成功，当前URL: {view_page.url}"

        # 阻断设备轮询请求，防止 session 被保活
        def _block_device(route):
            if "/api/device/" in route.request.url:
                route.abort()
            else:
                route.continue_()

        view_page.route("**/*", _block_device)

        # 导航到无设备轮询的静态页，让 session timer 开始倒计时
        view_page.goto(BASE_URL + "/#/systemSettings/dateTime")
        view_page.wait_for_load_state("networkidle")
        view_page.wait_for_timeout(1000)

        # 等待 1 分钟超时 + 15s 缓冲 = 75s
        view_page.wait_for_timeout(75_000)

        # 通过 API 验证 view 用户的 session 是否已过期（应返回 401）
        ss = view_page.evaluate("() => JSON.parse(sessionStorage.getItem('common') || '{}')")
        token = ss.get("authToken", "")
        resp = view_page.evaluate("""async (tok) => {
            try {
                const r = await fetch('/api/settings/ntpConfig?token=' + tok, {method: 'GET'});
                const b = await r.text();
                return {status: r.status, body: b.substring(0, 80)};
            } catch(e) { return {error: e.message}; }
        }""", token)

        status = resp.get("status", 0)
        assert status != 200, (
            f"1分钟无操作后 session 应过期，但 API 返回 {status}（200 表示仍有效）\n"
            f"token={token}\nbody={resp}"
        )
    finally:
        ctx.close()
        _restore_session_timeout(page)
