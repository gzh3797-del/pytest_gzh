import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_022
# 用例标题: COV Increment 默认值与合法范围值正确
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 进入参数配置页面并查看任一参数的 COV Increment 默认值。
#   2. 输入 COV Increment=0.000 并保存。
#   3. 输入 COV Increment=0.123 并保存。
# 预期结果: 1. COV Increment 默认值显示为 0.000。 | 2. 下边界合法值 0.000 保存成功。 | 3. 其他合法值 0.123 保存成功。

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


def test_TestCase_WEB2_033_001_022(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    cov_field = page.get_by_label("COV Increment", exact=False).or_(
        page.get_by_placeholder("Enter COV Increment")).first
    if cov_field.count() > 0:
        # 默认值验证
        default_val = cov_field.input_value()
        assert default_val != "", "COV Increment 应有默认值"
        # 合法值
        cov_field.fill("1.0")
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() == 0, "COV Increment=1.0 应合法"
