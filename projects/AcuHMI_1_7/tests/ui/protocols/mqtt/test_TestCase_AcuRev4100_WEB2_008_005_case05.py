import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_008_005_case05
# 用例标题: acucloud模块，检查字段"Devcies Selection"正确修改为"Devices Selection To Mapping红色星号"
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.AcuCloud服务端启动正常
# 测试步骤:
#   1、acucloud页面，检查Devices Selection To Mapping红色星号
# 预期结果: 1、Devices Selection To Mapping后面加红色星号 |  | 用例问题：4/28 Devices Selection To Mapping更改为Devices Selection

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

def test_TestCase_AcuRev4100_WEB2_008_005_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "Topic and Parameter Selection")
    page.wait_for_timeout(500)
    content = page.content()
    assert "Devcies" not in content, "旧拼写错误 'Devcies' 仍存在于AcuCloud/MQTT页面"
    assert "Devices Selection" in content, "'Devices Selection' 字段名应正确显示"
