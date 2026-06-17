import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_004_003
# 用例标题: 4100WEB2选择Virtual Device设备连接 AWS IoT Core 数据上报
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable | 3. 已配置正确的URL/Topic/Interval/证书文件
# 测试步骤:
#   1. Devices Selection To Mapping 勾选下挂的虚拟设备设备
#   2. MQTT Parameter Config页面Parameter Type 选择 AI，Parameter 选择 ALL
#   3. 连接AWS IoT
#   3. 使用 MQTT 客户端监控设备向云端上报的参数是否为配置的IOM设备参数
# 预期结果: 3. AWS IoT Core 收到正确的设备连接信息 | 4. MQTT客户端工具监控到设备向云端上报的数据为虚拟设备配置的参数，且参数值和参数计算公式计算的值一致

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


def test_TestCase_WEB2_AWS_004_003(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 选择 Virtual 并配置参数
    device_section = page.locator(".el-form-item, tr").filter(has_text="Device").first
    if device_section.count() > 0:
        device_row = page.locator("tr, .el-table__row").filter(has_text="Virtual").first
        if device_row.count() > 0:
            device_row.locator(".el-checkbox__inner").click()
            page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 运行 AWS IoT 客户端，验证收到该设备数据
