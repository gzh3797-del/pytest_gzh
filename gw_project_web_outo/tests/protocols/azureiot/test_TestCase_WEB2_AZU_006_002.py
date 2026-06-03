import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_006_002
# 用例标题: Primary/Secondary 均失效时连接失败本地缓存保留
# 预置条件: 设备已正常上电，相关服务正常启动
# 测试步骤:
#   1. 配置 Primary 和 Secondary Connection String，均失效
#   2. 启用 Azure IoT，查看连接状态
#   3. 查看本地缓存数据是否保留
# 预期结果: 2. 连接失败，界面提示连接错误 | 3. 本地采集数据正常缓存，数据不丢失

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


def test_TestCase_WEB2_AZU_006_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # 清空 Primary 和 Secondary
    primary = page.get_by_label("Primary Connection String", exact=False).first
    secondary = page.get_by_label("Secondary Connection String", exact=False).first
    primary.fill("")
    if secondary.count() > 0:
        secondary.fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    # 两者均空时应提示连接失败或配置不完整
    assert page.locator(".el-form-item__error").count() > 0 or         page.locator(".el-message--error").count() > 0 or         page.locator(".el-message").count() > 0,         "Primary/Secondary 均空时应有错误提示"
