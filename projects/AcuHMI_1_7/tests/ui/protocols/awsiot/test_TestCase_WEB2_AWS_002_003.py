import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_002_003
# 用例标题: Interval 参数校验与合法值校验
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable
# 测试步骤:
#   1. 进入 AWS IoT 配置页面，启用 Enable，配置合法 URL/Topic 及正确证书
#   2. Interval 下拉选择 1 seconds
#   3. 点击保存，使用MQTT客户端工具监控设备向云端上传数据的间隔
#   4. 重复步骤1-3，遍历Interval设置为10/30/60/90/120/180/240/300/480/600second
# 预期结果: 1. AWS IoT配置打开成功 | 2. Interval下拉列表选择1 second成功 | 3. 保存成功，使用MQTT客户端监控设备向云端上报参数信息的时间间隔为1s | 4. 设备向云端上报数据符合配置的Interval时间间隔

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


def test_TestCase_WEB2_AWS_002_003(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # Topic / Interval 参数校验
    topic_field = page.get_by_label("Topic", exact=False).or_(
        page.get_by_placeholder("Enter Topic")).first
    interval_field = page.get_by_label("Interval", exact=False).or_(
        page.get_by_placeholder("Enter Interval")).first

    if topic_field.count() > 0:
        topic_field.fill("aws/test/topic")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        topic_form = page.locator(".el-form-item").filter(has_text="Topic").first
        assert topic_form.locator(".el-form-item__error").count() == 0, "合法 Topic 应保存成功"

    # Interval is a dropdown (el-select readonly) — click to open and select an option
    interval_item = page.locator(".el-form-item").filter(has_text="Interval").first
    if interval_item.count() > 0:
        interval_item.locator(".el-select").click()
        page.wait_for_timeout(300)
        options = page.get_by_role("option").all()
        assert len(options) > 0, "Interval 下拉应有可选项"
        options[0].click()
        page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        interval_form = page.locator(".el-form-item").filter(has_text="Interval").first
        assert interval_form.locator(".el-form-item__error").count() == 0, "合法 Interval 应保存成功"
