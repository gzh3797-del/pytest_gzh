import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_005_004_case28
# 用例标题: 下拉设备选择AcuDio，选择参数DO Statue，查看客户端是否接收到推送的数据
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.MQTT服务端启动正常
# 测试步骤:
#   1.Communications->MQTT Device Parameter Config
#   2.下拉设备选择AcuDio，选择DO Statue
#   3.遍历参数DO Statue
#   4.查看客户端是否接收到推送的数据
#   5.选择全部的参数
#   6.查看客户端是否接收到推送的数据
# 预期结果: 4.客户端接收到推送数据 | 6.客户端接收到推送数据

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
    """Enable MQTT if disabled. Returns to original MQTT sub-page afterward."""
    current_url = page.url
    sub_url_map = {
        "credential": "User Credential",
        "ssl": "SSL",
        "testament": "Last Will and Testament",
        "deviceToPublish": "Topic and Parameter Selection",
    }
    sub_name_to_path = {v: k for k, v in sub_url_map.items()}
    original_sub = None
    for suffix, name in sub_url_map.items():
        if f"/protocols/mqtt/{suffix}" in current_url:
            original_sub = name
            break
    if "/protocols/mqtt/general" not in current_url:
        base = current_url.split("#")[0]
        page.goto(base + "#/protocols/mqtt/general")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(800)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)
    if original_sub:
        path = sub_name_to_path.get(original_sub, "")
        if path:
            base = page.url.split("#")[0]
            page.goto(base + f"#/protocols/mqtt/{path}")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)

def test_TestCase_AcuRev4100_WEB2_005_004_case28(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "Topic and Parameter Selection")
    _ensure_mqtt_enabled(page)

    # 选择 AcuDio 设备的 DO → DO Statue 参数
    device_row = page.locator("tr, .el-table__row").filter(has_text="AcuDio").first
    if device_row.count() == 0:
        pytest.skip("测试环境未找到 AcuDio 设备，请检查设备连接")

    device_row.locator(".el-checkbox__inner").click()
    page.wait_for_timeout(500)

    # 展开参数组 DO
    param_section = page.locator(".el-collapse-item, tr").filter(has_text="DO").first
    if param_section.count() > 0:
        param_section.click()
        page.wait_for_timeout(300)

    # 勾选参数 DO Statue
    param_row = page.locator("tr, .el-table__row").filter(has_text="DO Statue").first
    if param_row.count() > 0:
        param_row.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # TODO: 启动 MQTT 订阅客户端，等待 60s，断言收到含 DO Statue 参数的 JSON 消息
    # import paho.mqtt.client as mqtt
    # ... (需 MQTT broker 环境配合)
