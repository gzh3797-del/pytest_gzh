import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_040
# 用例标题: BACnet Port 非法边界外值被拦截
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 打开 BACnet/IP 页面。
#   2. 输入 BACnet Port=47807，点击 Save。
#   3. 输入 BACnet Port=49001，点击 Save。
#   4. 观察界面提示与保存结果。
# 预期结果: 1. BACnet Port 字段可编辑。 | 2. 下边界外值 47807 被判定为非法并禁止保存。 | 3. 上边界外值 49001 被判定为非法并禁止保存。 | 4. 界面给出明确错误提示。

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


def test_TestCase_WEB2_033_001_040(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # Enable BACnet/IP to make config fields visible
    bacnet_radio = page.locator(".el-form-item").filter(has_text="Enable").locator(".el-radio").filter(has_text="Enable")
    if bacnet_radio.count() > 0 and 'is-checked' not in (bacnet_radio.first.get_attribute("class") or ""):
        bacnet_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    f = page.get_by_label("BACnet Port", exact=False).or_(
        page.get_by_placeholder("Enter BACnet Port")).first

    # 合法值应保存成功
    for valid in ['47808', '48500', '49000']:
        f.fill(valid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() == 0, \
            f"BACnet Port={valid} 应保存成功"

    # 非法值应保存失败
    for invalid in ['47807', '49001', '0', '-1', 'abc']:
        f.fill(invalid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() > 0 or \
            page.locator(".el-message--error").count() > 0, \
            f"BACnet Port={invalid} 应保存失败"
