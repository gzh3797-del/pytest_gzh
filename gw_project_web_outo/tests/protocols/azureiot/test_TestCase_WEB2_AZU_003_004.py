import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_003_004
# 用例标题: Certificate 与 Key 不匹配时连接失败
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2. 进入 Azure IoT 配置页面，启用 Enable，开启 Enable SSL
# 测试步骤:
#   1. 上传不匹配的 Certificate 与 Key 文件（来自不同密钥对）
#   2. 保存配置，查看连接状态
# 预期结果: 1. 连接失败，界面提示认证错误；不向 Azure IoT Hub 发布任何数据

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


def test_TestCase_WEB2_AZU_003_004(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    ssl_item = page.locator(".el-form-item").filter(has_text="SSL")
    if ssl_item.count() > 0:
        ssl_item.locator(".el-radio, .el-switch, .el-checkbox").first.click()
        page.wait_for_timeout(300)

    assert page.locator("input[type=file]").count() > 0, "SSL证书上传控件应存在"
    # TODO: 上传格式错误证书/不匹配证书，断言系统拒绝或连接失败
