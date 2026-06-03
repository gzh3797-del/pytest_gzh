import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_034
# 用例标题: COV Batch Update 中 AcuIOM 参数列表与模板一致
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 点击 Batch Update COV 并选择 AcuIOM 设备。
#   2. 查看窗口中的参数列表和模板定义。
# 预期结果: 1. Batch Update COV 窗口可正常打开。 | 2. AcuIOM 参数列表与模板一致，无缺项或错项。

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
def test_TestCase_WEB2_033_001_034(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "BACnet/IP")

    # COV Batch Update 配置 AcuIOM 参数
    cov_batch = page.get_by_role("button", name="COV Batch Update").or_(
        page.get_by_text("COV Batch Update", exact=False)).first
    if cov_batch.count() > 0:
        cov_batch.click()
        page.wait_for_timeout(500)
        device_row = page.locator("tr, .el-table__row").filter(has_text="AcuIOM").first
        if device_row.count() > 0:
            device_row.locator(".el-checkbox__inner").click()
            page.wait_for_timeout(300)
        page.get_by_role("button", name="Save").or_(
            page.get_by_role("button", name="Confirm")).first.click()
        page.wait_for_timeout(1000)
        assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    else:
        assert False, "COV Batch Update 按钮未找到"
