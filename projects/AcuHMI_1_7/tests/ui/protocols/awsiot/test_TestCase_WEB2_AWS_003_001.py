import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_003_001
# 用例标题: 合法证书与密钥连接 AWS IoT Core
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable | 3. 已配置正确的URL/Topic/Interval
# 测试步骤:
#   1. 配置合法的Cert File 和 Key File
#   2. 保存配置，连接AWS IoT, 查看连接状态及 AWS IoT Core 数据接收情况
# 预期结果: 1. 证书和密钥文件上传成功 | 2. 设备正常连接 AWS IoT Core，连接状态显示成功；AWS IoT Core 收到设备上报数据

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


def test_TestCase_WEB2_AWS_003_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 验证证书上传控件存在
    cert_inputs = page.locator("input[type=file]")
    assert cert_inputs.count() > 0, "AWS IoT 页面应有证书上传控件"

    # 点击 Test Connection 并验证按钮存在
    test_btn = page.get_by_role("button", name="Test Connection").or_(
        page.get_by_role("button", name="Test")).first
    expect(test_btn).to_be_visible(timeout=5000), "Test Connection 按钮应可见"
    # TODO: 上传合法证书后点击 Test Connection 验证成功
