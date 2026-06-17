import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_006_001
# 用例标题: 禁用 AWS IoT 后停止发布重新启用后恢复
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable | 3. 已配置正确的URL/Topic/Interval/证书文件
# 测试步骤:
#   1. 设备已连接 AWS IoT，正常上报数据
#   2. 将 Enable 切换为 Disable，保存
#   3. 重新将 Enable 切换为 Enable，保存
# 预期结果: 2. 保存后不再向 AWS IoT Core 发布任何数据 | 3. 重新启用后，AWS IoT Core 可以正常收到设备上报数据

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


def test_TestCase_WEB2_AWS_006_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_item = page.locator(".el-form-item").filter(has_text="Enable").first

    # Disable
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Disable").click()
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # Enable
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 验证 Enable 后恢复向 AWS IoT 上报数据
