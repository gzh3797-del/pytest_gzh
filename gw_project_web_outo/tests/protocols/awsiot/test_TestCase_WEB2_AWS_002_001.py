import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_AWS_002_001
# 用例标题: URL 参数校验与合法值校验
# 预置条件: 1. 设备已正常上电，相关服务正常启动 | 2. 进入 AWS IoT 连接配置页面，启用 Enable
# 测试步骤:
#   1. 配置 URL 含非法字符（如 "TEST_url@iot.amazonaws.com"），点击保存
#   2. 配置 URL 长度超出 128 字符，点击保存
#   3. 配置合法 URL = "a1b2c3d4.iot.us-east-1.amazonaws.com"，点击保存
#   4. 配置URL为空，保存
# 预期结果: 1. 阻止保存，提示字符不合法（仅允许小写字母/数字/"-"".""/"） | 3. 阻止保存，提示长度超出范围（20~128） | 3. 保存成功，页面展示值与输入值一致 | 4. 阻止保存，提示URL不能为空

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


def test_TestCase_WEB2_AWS_002_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "AWS IoT")

    enable_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable").first
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    url_field = page.get_by_label("URL", exact=False).or_(
        page.get_by_placeholder("Enter URL").or_(
        page.get_by_placeholder("Enter Endpoint"))).first

    # 有效 URL - check URL field itself has no validation error
    url_field.fill("abcdefg.iot.us-east-1.amazonaws.com")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    url_form_item = page.locator(".el-form-item").filter(has_text="URL").or_(
        page.locator(".el-form-item").filter(has_text="Endpoint")).first
    assert url_form_item.locator(".el-form-item__error").count() == 0, "合法 URL 字段应无校验错误"

    # 非法 URL - URL field should show error
    url_field.fill("not valid url!!!")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    assert page.locator(".el-form-item__error").count() > 0 or \
        page.locator(".el-message--error").count() > 0, "非法 URL 应保存失败"
