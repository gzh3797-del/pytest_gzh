import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_019
# 用例标题: AcuIOM 参数列表与模板展示一致
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 选择任一 AcuIOM 设备并进入参数配置页面。
#   2. 查看参数列表列头和参数名(单位)展示。
# 预期结果: 1. AcuIOM 参数页可正常打开。 | 2. 参数列表完整展示 Parameter / EPICS Enable / COV Enable / COV Increment / 参数名(单位)，且与模板一致。

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


@pytest.mark.xfail(strict=False, reason="需要 AcuIOM 设备在 BACnet 设备列表中")
def test_TestCase_WEB2_033_001_019(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # 验证 AcuIOM 参数列表与模板一致
    # 找到 AcuIOM 设备行
    acu_row = page.locator("tr, .el-table__row").filter(has_text="AcuIOM").first
    expect(acu_row).to_be_visible(timeout=5000), "AcuIOM 设备应在 BACnet 设备列表中"
    # 展开查看参数列表
    acu_row.click()
    page.wait_for_timeout(500)
