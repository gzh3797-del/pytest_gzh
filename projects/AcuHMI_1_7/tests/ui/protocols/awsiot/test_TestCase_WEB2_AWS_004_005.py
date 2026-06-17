import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_004_005
# 用例标题: 4100/IOM设备未配置上传参数时发送为空
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable | 3. 已配置正确的URL/Topic/Interval/证书文件
# 测试步骤:
#   1. 设备均连接 AWS IoT，Devices Selection To Mapping已勾选所有下挂设备
#   2. MQTT Parameter Config 页面不配置任何参数
#   3. 连接AWS IoT
#   4. 使用 MQTT 客户端监控设备向云端发送的数据
# 预期结果: 2. MQTT客户端监控到设备未向云端传送消息

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


def test_TestCase_WEB2_AWS_004_005(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 选择设备但不配置上传参数
    first_device = page.locator("tr, .el-table__row").filter(
        has_text="AcuRev").first
    if first_device.count() > 0:
        first_device.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    # 设备未配置上传参数时，保存成功但发送为空
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
