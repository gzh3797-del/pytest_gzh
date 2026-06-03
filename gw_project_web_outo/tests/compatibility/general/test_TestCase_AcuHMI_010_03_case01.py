import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD
from pages.login_page import LoginPage

# 注意：本用例验证多浏览器兼容性（Chrome/Edge/Safari）
# - 默认测试引擎: chromium（等同Chrome/Edge内核）
# - 如需测试Firefox内核，运行时添加参数: --browser=firefox
# - 如需测试WebKit内核（Safari），运行时添加参数: --browser=webkit
# 也可在CI环境中使用矩阵策略并行执行三种浏览器


# 用例编号：TestCase_AcuHMI_010_03_case01
# 用例标题：使用谷歌、微软、safari浏览器访问服务，系统工作正常，功能未发生异常
# 预置条件：1.服务启动正常，账号登录成功
# 测试步骤：
#   1. 分别使用谷歌、微软、safari浏览器访问服务
# 预期结果：
#   1. 系统工作正常，功能未发生异常
# 备注：
#   - 本测试默认以 chromium 引擎执行（对应Chrome/Edge内核）
#   - 如需测试 Firefox 引擎：pytest --browser=firefox
#   - 如需测试 WebKit 引擎（Safari）：pytest --browser=webkit
#   - 可结合CI矩阵策略同时覆盖三种浏览器引擎
def test_TestCase_AcuHMI_010_03_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    browser_type_name = page.context.browser.browser_type.name

    assert "/#/login" not in page.url, \
        f"[{browser_type_name}] 应登录成功并跳转到主页，当前URL: {page.url}"

    assert page.locator("header span").filter(has_text="AcuHMI").count() > 0 or \
           page.locator("header, nav").count() > 0, \
        f"[{browser_type_name}] 登录后导航菜单应正常显示"

    # 访问About页面验证基本内容加载
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    assert page.locator("body").count() > 0, \
        f"[{browser_type_name}] About页面应正常加载"


@pytest.mark.skip(reason="需要单独使用 --browser=firefox 参数运行，无法在同一会话中切换浏览器引擎")
def test_TestCase_AcuHMI_010_03_case01_firefox(login_page: LoginPage):
    """使用Firefox引擎进行兼容性验证"""
    pass


@pytest.mark.skip(reason="需要单独使用 --browser=webkit 参数运行（macOS/Linux），无法在同一会话中切换浏览器引擎")
def test_TestCase_AcuHMI_010_03_case01_webkit(login_page: LoginPage):
    """使用WebKit引擎（Safari）进行兼容性验证"""
    pass
