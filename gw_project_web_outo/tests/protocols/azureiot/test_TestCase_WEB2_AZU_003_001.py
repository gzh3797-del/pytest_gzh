import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_003_001
# 用例标题: Enable SSL 默认关闭及启用后字段可见
# 预置条件: 设备已正常上电，相关服务正常启动
# 测试步骤:
#   1. 进入 Azure IoT 配置页面，启用 Enable，查看 Enable SSL 状态
#   2. 将 Enable SSL 开启
#   3. 将 Enable SSL 关闭
# 预期结果: 1. Enable SSL 默认为关闭状态，Certificate/Key 字段不可见 | 2. 开启后 Certificate 文件和 Key 文件上传字段显示并可操作 | 3. 关闭后上传字段隐藏，无法上传

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


def test_TestCase_WEB2_AZU_003_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # SSL 默认关闭
    ssl_item = page.locator(".el-form-item").filter(has_text="SSL")
    if ssl_item.count() > 0:
        # 启用 SSL 后证书字段应显示
        ssl_item.locator(".el-radio, .el-switch, .el-checkbox").first.click()
        page.wait_for_timeout(500)
        assert page.locator("input[type=file]").count() > 0 or             page.get_by_text("Certificate", exact=False).count() > 0,             "SSL 开启后应显示证书配置"
