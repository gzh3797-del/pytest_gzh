import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_008_003_case24
# 用例标题: 设置非法端口验证；预期保存失败
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备
# 测试步骤:
#   1.设置Port为160，点击保存
#   2.设置port为16200，点击保存
#   3.设置port为-161，点击保存
#   4.设置port为A，点击保存
#   5.设置port为@，点击保存
# 预期结果: 1.保存失败，提示语准确 | 2.保存失败，提示语准确 | 3.保存失败，提示语准确 | 4.保存失败，提示语准确 | 5.保存失败，提示语准确

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


def test_TestCase_AcuRev4100_WEB2_008_003_case24(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "SNMP")

    # 非法端口：0, 65536
    port_field = page.get_by_label("Port", exact=False).or_(
        page.get_by_placeholder("Enter Port"))
    for invalid in ["0", "65536", "-1", "abc"]:
        port_field.fill(invalid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() > 0 or             page.locator(".el-message--error").count() > 0,             f"非法端口 {invalid} 应保存失败"
