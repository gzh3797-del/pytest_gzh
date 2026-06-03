import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_006_001
# 用例标题: Primary 失效时自动切换 Secondary 继续上报
# 预置条件: 设备已正常上电，相关服务正常启动
# 测试步骤:
#   1. 配置合法 Primary 和 Secondary Connection String，已连接 Azure IoT Hub
#   2. 使 Primary Connection String 失效（如密钥过期）
#   3. 查看连接状态及数据上报情况
# 预期结果: 3. 系统自动切换至 Secondary Connection String，继续向 Azure IoT Hub 正常上报数据，无数据中断

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


def test_TestCase_WEB2_AZU_006_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_item = page.locator(".el-form-item").filter(has_text="Enable").first

    # 配置 Primary，清空模拟失效，Secondary 接管
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.wait_for_timeout(300)

    valid_cs = "HostName=myhub.azure-devices.net;DeviceId=mydevice;SharedAccessKey=abc123def456=="
    primary = page.get_by_label("Primary Connection String", exact=False).first
    secondary = page.get_by_label("Secondary Connection String", exact=False).first
    primary.fill(valid_cs)
    if secondary.count() > 0:
        secondary.fill(valid_cs.replace("mydevice", "dev2"))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 清空 Primary
    primary.fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 验证系统自动切换到 Secondary
