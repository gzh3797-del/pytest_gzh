import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_004_001
# 用例标题: 选择Modbus设备发布数据按间隔上报
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2. 进入 Azure IoT 配置页面，启用 Enable
# 测试步骤:
#   1. 配置合法 Connection String 和 Interval
#   2. Select Devices to Publish 选择 Modbus设备
#   3. MQTT Parameter Config配置设备上报参数
#   4. 连接Azure IoT，查看设备上报的数据参数和数值是否正确
# 预期结果: 3. 配置成功 | 4. 连接成功，Azure IoT Hub 按配置的 Interval 间隔正常收到Modbus设备参数数据，上报的参数个数和配置的一致，且数值正确

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


def test_TestCase_WEB2_AZU_004_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 选择 AcuRev 设备
    device_row = page.locator("tr, .el-table__row").filter(has_text="AcuRev").first
    if device_row.count() > 0:
        device_row.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(1000)
        assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    else:
        assert False, "测试环境未找到 AcuRev 设备"
    # TODO: 运行 Azure IoT 客户端，验证收到该设备数据
