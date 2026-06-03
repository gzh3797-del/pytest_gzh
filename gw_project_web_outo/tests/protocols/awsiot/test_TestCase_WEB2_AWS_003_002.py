import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_003_002
# 用例标题: 非法证书或错误密钥连接失败
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable | 3. 已配置正确的URL/Topic/Interval
# 测试步骤:
#   1. 上传正确的证书文件，错误的密钥文件，查看连接是否成功
#   2. 上传错误的证书文件，正确的密钥文件，查看连接是否成功
#   3. 上传错误的证书文件，错误的密钥文件，查看连接是否成功
#   4. 上传.txt/损坏的证书或密钥文件，查看连接是否成功
# 预期结果: 1. 错误的密钥文件无法连接AWS IoT | 2. 错误的证书文件无法连接AWS IoT | 3. 错误的证书和错误的密钥无法连接AWS IoT平台 | 4. 错误格式的证书或密钥文件无法连接AWS IoT平台

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


def test_TestCase_WEB2_AWS_003_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    assert page.locator("input[type=file]").count() > 0, "应有证书上传控件"
    # TODO: 上传格式错误的证书，点击 Test Connection，断言显示失败提示
    test_btn = page.get_by_role("button", name="Test Connection").or_(
        page.get_by_role("button", name="Test")).first
    expect(test_btn).to_be_visible(timeout=5000)
