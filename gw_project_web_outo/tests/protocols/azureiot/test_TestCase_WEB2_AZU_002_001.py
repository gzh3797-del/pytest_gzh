import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_002_001
# 用例标题: Connection String 合法值配置保存连接
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2. 进入 Azure IoT 配置页面，启用 Enable
# 测试步骤:
#   1. 配置合法 Primary Connection String（含 HostName/DeviceId/SharedAccessKey）和Primary Connection String（含 HostName/DeviceId/SharedAccessKey）
#   2. 点击Test Connection，查看连接状态
# 预期结果: 1. Primary Connection String 字段接受输入 | 2. 保存成功，展示值正确；设备正常连接 Azure IoT Hub

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


def test_TestCase_WEB2_AZU_002_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    conn_str_field = page.get_by_label("Primary Connection String", exact=False).or_(
        page.get_by_placeholder("Enter Connection String")).first
    # 合法 Connection String 格式
    valid_cs = "HostName=myhub.azure-devices.net;DeviceId=mydevice;SharedAccessKey=abc123def456=="
    conn_str_field.fill(valid_cs)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-form-item__error").count() == 0,         "合法 Connection String 应保存成功"
