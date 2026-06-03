import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_001_001
# 用例标题: AWS IoT 配置开启与关闭
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable
# 测试步骤:
#   1. 点击通信协议 -> AWS IoT 配置页面；
#   2. 默认AWS IoT Enable为Disable,查看连接参数是否显示；
#   3. 点击Enable，打开AWS IoT配置，查看URL/Topic/Interval/Cert File/key File，Test Connection是否显示且可编辑
# 预期结果: 1. 页面正常打开， | 2. Enable 状态默认为 Disable；URL/Topic/Interval/Cert File/Key File 字段及Test Connection 按钮均不显示

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


def test_TestCase_WEB2_AWS_001_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")
    page.wait_for_timeout(500)

    enable_item = page.locator(".el-form-item").filter(has_text="Enable").first
    expect(enable_item).to_be_visible(timeout=5000)

    # 切换 Enable
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Enable").click()
    page.wait_for_timeout(300)
    # 配置字段应显示
    assert page.locator(".el-form-item").count() > 2, "Enable 后配置字段应显示"

    # 切换 Disable
    enable_item.locator(".el-radio, .el-switch").filter(has_text="Disable").click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
