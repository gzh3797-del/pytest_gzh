import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_008_004_case33
# 用例标题: MQTT模块，需求文档要求basic topic的长度限制在128，设置长度为129
# 预置条件: 1、服务启动正常，账号登录成功 | 2、MQTT服务端启动正常
# 测试步骤:
#   1、进入 MQTT > Topic and Parameter Selection 页面
#   2、在 Base Topic 字段输入 129 个字符
#   3、点击 Save
# 预期结果: 保存失败，页面显示校验错误


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
    """Enable MQTT if it is currently disabled."""
    base = page.url.split("#")[0]
    page.goto(base + "#/protocols/mqtt/general")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    if enable_item.count() > 0:
        enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
        if "is-checked" not in (enable_radio.get_attribute("class") or ""):
            enable_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(500)
            page.get_by_role("button", name="Save").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(800)


def test_TestCase_AcuRev4100_WEB2_008_004_case33(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 确保 MQTT 已启用
    _nav_protocol(page, "MQTT")
    _ensure_mqtt_enabled(page)

    # 进入 Topic and Parameter Selection
    _nav_protocol(page, "MQTT", "Topic and Parameter Selection")

    # 定位 Base Topic 输入框（placeholder="Enter Base Topic"）
    base_topic_input = page.locator(".el-form-item").filter(has_text="Base Topic").first.locator("input").first
    base_topic_input.wait_for(state="visible", timeout=5000)

    # 输入 129 字符（超过限制 128）
    long_topic = "a" * 129
    base_topic_input.fill(long_topic)
    page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)

    # 验证保存失败：出现字段校验错误或错误消息
    field_error = page.locator(".el-form-item__error").count()
    msg_error = page.locator(".el-message--error").count()
    assert field_error > 0 or msg_error > 0, \
        f"Base Topic 长度129应保存失败（最大128），但未出现错误提示（field_error={field_error}, msg_error={msg_error}）"
