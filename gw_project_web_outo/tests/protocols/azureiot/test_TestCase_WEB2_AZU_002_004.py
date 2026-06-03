import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AZU_002_004
# 用例标题: Interval = 10s（最小值）配置保存及上报
# 预置条件: 1.设备已正常上电，相关服务正常启动 | 2. 进入 Azure IoT 配置页面，启用 Enable，配置合法 Connection String
# 测试步骤:
#   1. 配置 Interval = 10 seconds（最小值）
#   2. 点击保存，查看数据上报间隔
#   3. 重复步骤1-3步骤，遍历Interval为10/100/300/600
# 预期结果: 1. 字段接受输入 | 2. 保存成功，展示值为 10 seconds；Azure IoT Hub 按约 10 秒间隔收到数据

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


def test_TestCase_WEB2_AZU_002_004(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "Azure IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # Interval is a dropdown (el-select readonly) — click to open and select an option
    interval_item = page.locator(".el-form-item").filter(has_text="Interval").first
    interval_item.locator(".el-select").click()
    page.wait_for_timeout(300)
    options = page.get_by_role("option").all()
    assert len(options) > 0, "Interval 下拉应有可选项"
    # Select the smallest available option (10s minimum per spec)
    options[0].click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    interval_form = page.locator(".el-form-item").filter(has_text="Interval").first
    assert interval_form.locator(".el-form-item__error").count() == 0, "Interval 最小选项应保存成功"
