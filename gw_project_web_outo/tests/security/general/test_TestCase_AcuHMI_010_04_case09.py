import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL, DEFAULT_USERNAME
from pages.login_page import LoginPage

# 敏感信息关键词：这些内容不应出现在页面错误信息中
_SENSITIVE_PATTERNS = [
    "stack trace",
    "stacktrace",
    "Traceback",
    "at line",
    "File \"",
    "SyntaxError",
    "ReferenceError",
    "NullPointerException",
    "java.lang",
    "org.springframework",
    "/etc/",
    "/var/",
    "/home/",
    "C:\\",
    "node_modules",
    "webpack",
    "SQL",
    "SELECT ",
    "INSERT ",
    "UPDATE ",
    "version:",
    "v1.",
    "v2.",
    "v3.",
    "nginx/",
    "Apache/",
    "Express",
]


def _page_text_contains_sensitive(page) -> list:
    """检查当前页面可见文本是否包含敏感信息关键词，返回匹配列表"""
    body_text = page.locator("body").inner_text().lower()
    found = []
    for pattern in _SENSITIVE_PATTERNS:
        if pattern.lower() in body_text:
            found.append(pattern)
    return found


# 用例编号：TestCase_AcuHMI_010_04_case09
# 用例标题：网页中所有的系统错误信息刷新，均不存在敏感信息
# 预置条件：1.服务启动正常，账号登录成功
# 测试步骤：
#   1. 网站是否泄露敏感信息通过错误消息
# 预期结果：
#   1. 不泄露敏感信息通过错误消息（错误提示不含堆栈/路径/版本号等系统内部信息）
def test_TestCase_AcuHMI_010_04_case09(login_page: LoginPage):
    page = login_page.page

    # --- 场景1: 访问不存在的页面URL，触发404类错误 ---
    page.goto(BASE_URL + "/#/nonexistent/page/that/does/not/exist")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    found = _page_text_contains_sensitive(page)
    assert not found, \
        f"访问不存在页面时，错误信息不应泄露敏感内容，检测到: {found}"

    # --- 场景2: 尝试未授权访问API ---
    api_responses = []

    def _capture_api_response(response):
        if "/api/" in response.url:
            try:
                body = response.text()
                api_responses.append({"url": response.url, "status": response.status, "body": body})
            except Exception:
                pass

    page.on("response", _capture_api_response)

    # 使用错误的token访问API，触发401错误响应
    result = page.evaluate("""async () => {
        try {
            const r = await fetch('/api/settings/ntpConfig', {
                method: 'GET',
                headers: {'Authorization': 'Bearer invalid_token_12345'}
            });
            return {status: r.status, body: (await r.text()).substring(0, 500)};
        } catch(e) { return {error: e.message}; }
    }""")
    api_body = result.get("body", "").lower()
    for pattern in _SENSITIVE_PATTERNS:
        assert pattern.lower() not in api_body, \
            f"API错误响应体不应包含敏感信息关键词 '{pattern}'，响应体: {api_body[:200]}"

    # --- 场景3: 登录时使用错误密码，触发认证错误 ---
    login_page.open()
    page.wait_for_load_state("networkidle")
    page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
    page.get_by_role("textbox", name="Enter Password").fill("WrongPassword!!")
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_timeout(1000)

    found = _page_text_contains_sensitive(page)
    assert not found, \
        f"登录失败时，错误提示不应包含敏感系统内部信息，检测到: {found}"

    # --- 场景4: 登录后正常操作，验证系统信息不泄露 ---
    login_page.open()
    login_page.login()
    page.goto(BASE_URL + "/#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    found = _page_text_contains_sensitive(page)
    # 过滤可能的误报（如版本号在About页面是正常显示的）
    # 只关注错误消息区域
    error_msgs = page.locator(".el-message--error, .el-form-item__error, .error-msg, [class*='error']")
    if error_msgs.count() > 0:
        error_text = error_msgs.all_inner_texts()
        for text in error_text:
            for pattern in _SENSITIVE_PATTERNS:
                assert pattern.lower() not in text.lower(), \
                    f"系统错误信息不应泄露敏感信息: '{pattern}'，错误文本: '{text}'"
