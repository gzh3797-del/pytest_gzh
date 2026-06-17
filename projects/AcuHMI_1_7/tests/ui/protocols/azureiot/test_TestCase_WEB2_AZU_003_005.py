import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_003_005
# 用例标题: SSL 开启时加密连接数据正常上报
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2.进入 Azure IoT 配置页面，启用 Enable，开启 Enable SSL
# 测试步骤:
#   1. 上传合法 X509 Certificate 和 Key，保存
#   2. 查看 Azure IoT Hub 接收加密数据情况
# 预期结果: 2. 使用 X509 证书加密连接 Azure IoT Hub 成功；Azure IoT Hub 正常接收加密上报数据

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


def test_TestCase_WEB2_AZU_003_005(login_page: LoginPage):
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
    # TODO: 上传合法 X509 证书，点击 Test Connection，断言成功
