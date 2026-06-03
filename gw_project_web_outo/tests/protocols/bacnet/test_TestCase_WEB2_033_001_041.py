import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_041
# 用例标题: Network Number 非法边界外值被拦截
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 打开 BACnet/IP 页面。
#   2. 输入 Network Number=0，点击 Save。
#   3. 输入 Network Number=65535，点击 Save。
#   4. 观察界面提示与保存结果。
# 预期结果: 1. Network Number 字段可编辑。 | 2. 下边界外值 0 被判定为非法并禁止保存。 | 3. 上边界外值 65535 被判定为非法并禁止保存。 | 4. 界面给出明确错误提示。

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


def test_TestCase_WEB2_033_001_041(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    f = page.get_by_label("Network Number", exact=False).or_(
        page.get_by_placeholder("Enter Network Number")).first

    # 合法值应保存成功
    for valid in ['1', '100', '65534']:
        f.fill(valid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() == 0, \
            f"Network Number={valid} 应保存成功"

    # 非法值应保存失败
    for invalid in ['0', '65535', '-1']:
        f.fill(invalid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() > 0 or \
            page.locator(".el-message--error").count() > 0, \
            f"Network Number={invalid} 应保存失败"
