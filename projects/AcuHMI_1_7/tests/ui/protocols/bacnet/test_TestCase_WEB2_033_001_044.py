import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_044
# 用例标题: Time To Live 非法边界外值被拦截
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 启用 Enable Foreign Device Function。
#   2. 输入 Time To Live=4 并保存。
#   3. 输入 Time To Live=1441 并保存。
#   4. 观察界面提示与保存结果。
# 预期结果: 1. Foreign Device 相关字段可编辑。 | 2. 下边界外值 4 被判定为非法并禁止保存。 | 3. 上边界外值 1441 被判定为非法并禁止保存。 | 4. 界面给出明确错误提示。

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


def test_TestCase_WEB2_033_001_044(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # Enable Foreign Device Function
    foreign_radio = page.locator(".el-form-item").filter(
        has_text="Foreign Device").locator(".el-radio").filter(has_text="Enable")
    if foreign_radio.count() > 0 and 'is-checked' not in (foreign_radio.get_attribute("class") or ""):
        foreign_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(300)

    f = page.get_by_label("Time To Live", exact=False).or_(
        page.get_by_placeholder("Enter Time To Live")).first

    # 合法值应保存成功（TTL 合法范围 5-1440）
    for valid in ['5', '60', '1440']:
        f.fill(valid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() == 0, \
            f"Time To Live={valid} 应保存成功"

    # 非法值应保存失败
    for invalid in ['4', '1441', '0', '-1']:
        f.fill(invalid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() > 0 or \
            page.locator(".el-message--error").count() > 0, \
            f"Time To Live={invalid} 应保存失败"
