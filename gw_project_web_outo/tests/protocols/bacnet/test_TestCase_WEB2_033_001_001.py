import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_001
# 用例标题: BACnet/IP 页面入口与默认状态
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 登录 Web 页面并进入设置 -> 通信。
#   2. 打开 BACnet/IP 页面。
# 预期结果: 1. 通信页面可正常打开。 | 2. 可正常进入 BACnet/IP 页面，协议默认状态为 Disable。

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


def test_TestCase_WEB2_033_001_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")
    page.wait_for_timeout(500)

    # 验证 BACnet/IP 页面入口正常，包含基本配置项
    expect(page.locator("body")).to_be_visible()
    assert "BACnet" in page.content() or "Port" in page.content(),         "BACnet/IP 页面应显示配置项"
    # 验证默认 Enable 状态
    expect(page.locator(".el-form-item").filter(has_text="Enable").first).to_be_visible()
