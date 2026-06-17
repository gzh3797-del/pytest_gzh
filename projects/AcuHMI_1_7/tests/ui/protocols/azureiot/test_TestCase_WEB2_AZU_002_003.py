import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_002_003
# 用例标题: Secondary Connection String 配置及 Primary 失效切换
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2. 进入 Azure IoT 配置页面，启用 Enable
# 测试步骤:
#   1. 配置合法 Secondary Connection String，Primary 留空，点击保存
#   2. 使 Primary Connection String 失效，观察连接行为
# 预期结果: 1. 保存成功，Secondary 字段为空时允许保存 | 2. Primary 失效后系统自动切换至 Secondary Connection String，Azure IoT Hub 继续正常收到数据

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


def test_TestCase_WEB2_AZU_002_003(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 配置 Primary 和 Secondary Connection String
    valid_cs = "HostName=myhub.azure-devices.net;DeviceId=mydevice;SharedAccessKey=abc123def456=="
    primary = page.get_by_label("Primary Connection String", exact=False).first
    secondary = page.get_by_label("Secondary Connection String", exact=False).first
    primary.fill(valid_cs)
    if secondary.count() > 0:
        secondary.fill(valid_cs.replace("mydevice", "mydevice2"))
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 清空 Primary 模拟失效，验证系统切换到 Secondary
    primary.fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 验证系统切换到 Secondary 连接并继续推送数据
