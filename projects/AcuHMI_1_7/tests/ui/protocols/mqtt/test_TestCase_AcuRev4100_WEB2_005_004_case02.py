import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_005_004_case02
# 用例标题: Broker Port 设置端口号：有效端口号，无效端口号，非法端口号
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.MQTT服务端启动正常
# 测试步骤:
#   1.MQTT Enable启用
#   2.Broker Port设置端口号为1883
#   3.点击save，查看客户端是否接收到推送消息
#   4.Broker Port设置端口号为aaaa
#   5.点击save
#   6.Broker Port设置端口号为0
#   7.点击save
#   8.Broker Port设置端口号为65536
#   9.点击save
# 预期结果: 3.提示保存成功，客户端可以正常接收推送的数据 | 5.提示保存失败 | 7.提示保存成功 | 9.提示保存失败

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

def test_TestCase_AcuRev4100_WEB2_005_004_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "General")
    _ensure_mqtt_enabled(page)

    port_field = page.get_by_label("Broker Port", exact=False).or_(
        page.get_by_placeholder("Enter Broker Port"))

    # 有效端口(1883)
    port_field.fill("1883")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 有效端口0（期望保存成功）
    port_field.fill("0")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() == 0, "端口0应保存成功（测试规格：step7提示保存成功）"

    # 非法端口(abc)：期望失败
    port_field.fill("abc")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() > 0 or \
        page.locator(".el-message--error").count() > 0, "非法端口应保存失败"

    # 越界端口65536：期望失败
    port_field.fill("65536")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() > 0 or \
        page.locator(".el-message--error").count() > 0, "端口65536应保存失败"
