import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_005_004_case01
# 用例标题: Broker Address设置地址：有效地址，无效地址，非法地址
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.MQTT服务端启动正常
# 测试步骤:
#   1.MQTT Enable启用
#   2.Broker Address内输入有效域名地址
#   3.点击save，查看是否提示保存成功/无变化
#   4.Broker Address内输入非法字符
#   5.点击save，查看是否提示保存失败
# 预期结果: 有效地址保存成功；非法地址显示校验错误

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


def _ensure_mqtt_enabled(page):
    """Enable MQTT if currently disabled so config fields are visible."""
    if "/protocols/mqtt/general" not in page.url:
        page.get_by_role("menuitem", name="MQTT").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(800)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)


def _fill_mqtt_defaults(page):
    """Fill required MQTT fields with defaults so individual field tests can work."""
    # Fill broker if empty
    broker_field = page.get_by_placeholder("Enter Broker Address")
    if not broker_field.input_value():
        broker_field.fill("test.broker.com")
    # Generate client ID if empty
    client_id_field = page.get_by_placeholder("Enter Client ID")
    if not client_id_field.input_value():
        gen_btn = page.get_by_role("button", name="Generate Client ID")
        if gen_btn.count() > 0:
            gen_btn.click()
            page.wait_for_timeout(300)


def test_TestCase_AcuRev4100_WEB2_005_004_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "General")
    _ensure_mqtt_enabled(page)
    _fill_mqtt_defaults(page)

    broker_field = page.get_by_placeholder("Enter Broker Address")

    # 有效域名地址：保存后不显示表单校验错误
    broker_field.fill("test.broker.com")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() == 0, "有效域名地址应保存成功，不应显示校验错误"
    # Check message appeared (success or "no change", both acceptable)
    msg_count = page.locator(".el-message").count()
    assert msg_count > 0, "保存后应显示消息提示"
    assert page.locator(".el-message--error").count() == 0, "有效地址不应显示错误消息"
    page.wait_for_timeout(3000)  # let toast disappear

    # 非法字符地址：显示校验错误
    broker_field.fill("!@#$%^invalid!!!")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() > 0, "非法地址应显示校验错误"
