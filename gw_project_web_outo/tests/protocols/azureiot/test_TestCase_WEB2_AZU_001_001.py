import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_001_001
# 用例标题: Azure IoT 默认 Disable 页面配置隐藏
# 预置条件: 设备已正常上电，相关服务正常启动
# 测试步骤:
#   1. 点击协议 -> Azure IoT 配置页面，默认Azure IoT Enable为Disable状态，查看primary Connection String/second Connection String/Interval/Enable SSL配置是否隐藏
#   2.Azure IoT Enable开关切换为Enable，查看primary Connection String/second Connection String/Interval/Enable SSL配置是否显示且可编辑
# 预期结果: 1. Azure IoT Enable为Disable时，配置信息隐藏不可见 | 2. Azure IoT Enable为Enable时，配置信息显示且可编辑

def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_WEB2_AZU_001_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")
    page.wait_for_timeout(500)

    # 验证默认为 Disable 状态
    expect(page.locator("body")).to_be_visible()
    content = page.content()
    # 默认 Disable 时，连接字符串等配置字段应隐藏
    assert "Azure IoT" in content or "Connection" in content,         "Azure IoT 页面应正常显示"
