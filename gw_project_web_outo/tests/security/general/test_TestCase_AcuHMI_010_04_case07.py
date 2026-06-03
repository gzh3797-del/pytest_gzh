import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

# 需要通过HTTPS访问的页面URL路径列表
_PAGES_TO_CHECK = [
    BASE_URL + "/#/login",
    BASE_URL + "/#/systemSettings/dateTime",
    BASE_URL + "/#/userManagement/general",
    BASE_URL + "/#/maintenance/eventLog",
    BASE_URL + "/#/diagnostics/network",
]


# 用例编号：TestCase_AcuHMI_010_04_case07
# 用例标题：访问网站，以https安全的方式访问网页
# 预置条件：1.服务启动正常，账号登录成功
# 测试步骤：
#   1. 网站所有敏感数据传输是否使用https
# 预期结果：
#   1. 是（所有页面URL以https://开头）
def test_TestCase_AcuHMI_010_04_case07(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 验证BASE_URL本身以https开头
    assert BASE_URL.startswith("https://"), \
        f"系统BASE_URL应以https://开头，当前为: '{BASE_URL}'"

    # 逐页检查URL以https开头
    for url in _PAGES_TO_CHECK:
        page.goto(url)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        current_url = page.url
        assert current_url.startswith("https://"), \
            f"访问页面 '{url}' 后当前URL应以https://开头，实际为: '{current_url}'"

    # 验证登录请求等API接口也通过HTTPS发送（拦截请求检查协议）
    api_requests = []

    def _record_request(request):
        if "/api/" in request.url:
            api_requests.append(request.url)

    page.on("request", _record_request)

    # 触发一次API请求（访问系统设置页面）
    page.goto(BASE_URL + "/#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    for req_url in api_requests:
        assert req_url.startswith("https://") or req_url.startswith("/"), \
            f"API请求应通过HTTPS发送，实际URL: '{req_url}'"
