import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_007_002
# 用例标题: Azure IoT 与 AWS IoT 同时启用互不干扰
# 预置条件: 设备已正常上电，相关服务正常启动
# 测试步骤:
#   1. 同时启用 Azure IoT 和 AWS IoT，均配置合法参数并保存
#   2. 分别查看两侧云端接收数据情况
# 预期结果: 2. Azure IoT Hub 和 AWS IoT Core 各自独立正常接收设备上报数据；双侧数据互不干扰

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


def test_TestCase_WEB2_AZU_007_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 验证 Azure IoT 与 AWS IoT 同时启用互不干扰
    # 先启用 AWS IoT
    _nav_protocol(page, "AWS IoT")
    aws_enable = page.locator(".el-form-item").filter(has_text="Enable").first
    aws_enable.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 再启用 Azure IoT
    _nav_protocol(page, "Azure IoT")
    azu_enable = page.locator(".el-form-item").filter(has_text="Enable").first
    azu_enable.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 验证两者同时启用时数据推送互不干扰
