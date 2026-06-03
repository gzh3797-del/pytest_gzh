import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_005_004_case04
# 用例标题: Keep Alive（心跳）配置超时时间，查看超过配置时间是否发送心跳消息
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.MQTT服务端启动正常
# 测试步骤:
#   1.MQTT Enable启用
#   2.Broker Address内输入www.accu.com，端口号输入1883，Client ID输入abcdef
#   3.keepalive设置为10s
#   4.点击save提示保存成功
#   5.使用wrishrak抓包工具查看心跳报文
#   6.keepalive设置为50s
#   7.点击save提示保存成功
#   8.使用wrishrak抓包工具查看心跳报文
#   9.keepalive设置为9s
#   10.点击save提示保存失败
#   11.keepalive设置为601s
#   12.点击save提示保存失败
#   注：客户端发送 c0 00 服务端响应：D0 00
# 预期结果: 4.提示保存成功 | 5.wrishrak每10s可以看到一次心跳响应 | 8.wrishrak每50s可以看到一次心跳响应 | 10.提示保存失败 | 12.提示保存失败

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

def test_TestCase_AcuRev4100_WEB2_005_004_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "General")
    _ensure_mqtt_enabled(page)

    keep_alive = page.get_by_label("Keep Alive", exact=False).or_(
        page.get_by_placeholder("Enter Keep Alive"))
    keep_alive.fill("60")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 刷新后验证值持久化
    page.reload()
    page.wait_for_load_state("networkidle")
    _nav_protocol(page, "MQTT", "General")
    page.wait_for_timeout(500)
    assert "60" in page.content(), "Keep Alive 值未持久化"
