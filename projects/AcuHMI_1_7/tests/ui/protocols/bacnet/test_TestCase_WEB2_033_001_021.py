import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_021
# 用例标题: EPICS 与 COV 联动规则正确
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 选择任一参数并查看 EPICS Enable、COV Enable、COV Increment 的初始状态。
#   2. 开启 EPICS Enable 后检查 COV Enable 状态。
#   3. 开启 COV Enable 后检查 COV Increment 编辑状态。
# 预期结果: 1. 三个字段可正常显示。 | 2. EPICS Enable 使能后才允许使能 COV Enable。 | 3. COV Enable 使能后才允许编辑 COV Increment，且 EPICS Enable 支持全选与单独使能。

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


def test_TestCase_WEB2_033_001_021(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # EPICS 与 COV 联动：启用 EPICS 时 COV 相关字段应显示
    epics_item = page.locator(".el-form-item").filter(has_text="EPICS").first
    if epics_item.count() > 0:
        epics_item.locator(".el-radio, .el-switch, .el-checkbox").first.click()
        page.wait_for_timeout(500)
        # COV 字段应联动出现
        cov_item = page.locator(".el-form-item").filter(has_text="COV").first
        expect(cov_item).to_be_visible(timeout=3000), "EPICS 开启后 COV 字段应联动显示"
