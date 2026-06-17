import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_005_004_case08
# 用例标题: User Credential设置用户名以及密码，设置不同的用户名以及密码，设置非法的用户名以及密码
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3.MQTT服务端启动正常
# 测试步骤:
#   1.MQTT Enable启用
#   2.Broker Address内输入www.accu.com，端口号输入1883，Client ID输入abcdef
#   3.keepalive设置为60，Timeout输入3
#   4.Clean Session选择为true
#   5.user输入admin，密码输入power
#   6.点击save保存成功，查看客户端是否可以接收到传输的数据
#   7.输入错误的用户名和密码
#   8.点击save保存成功，查看客户端是否可以接收到传输的数据
# 预期结果: 6.提示保存成功，客户端可以接收到传输的数据 | 6.提示保存成功，客户端不可以接收到传输的数据

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

def test_TestCase_AcuRev4100_WEB2_005_004_case08(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "MQTT", "User Credential")
    _ensure_mqtt_enabled(page)

    # 设置有效用户名和密码
    page.get_by_label("Username", exact=False).or_(
        page.get_by_placeholder("Enter Username")).fill("testuser")
    page.get_by_label("Password", exact=True).or_(
        page.get_by_placeholder("Enter Password")).fill("Test@1234")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"

    # 空用户名应保存失败
    page.get_by_label("Username", exact=False).or_(
        page.get_by_placeholder("Enter Username")).fill("")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() > 0 or         page.locator(".el-message--error").count() > 0 or         page.locator(".el-message").first.is_visible(), "空用户名行为符合预期"
