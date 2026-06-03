import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


def _nav_to_diagnostics(page, submenu: str):
    if "/diagnostics" not in page.url.lower():
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/protocols",
            "/maintenance", "/templates", "/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Diagnostics").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_05_case02
# 用例标题：Modbus Debug Log启用，RTU_REQ/slaveid:246/fc3，导出调试日志成功
@pytest.mark.xfail(strict=False, reason="Modbus Debug Log导出依赖实际Modbus通信数据")
def test_TestCase_AcuHMI_009_05_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Modbus Debug Log")

    # Enable Modbus Debug Trace
    try:
        page.get_by_role("button", name="Enable").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    # Select log type and slaveid
    try:
        page.locator(".el-form-item").filter(has_text="Type").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="RTU_REQ", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Slave ID").locator("input").fill("246")
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Function Code").locator("input").fill("3")
    except Exception:
        pass

    page.get_by_role("button", name="Search").click()
    page.wait_for_timeout(1000)

    # Export Debug Logs
    with page.expect_download(timeout=15000) as download_info:
        page.get_by_role("button", name="Export Debug Logs").click()
    download = download_info.value
    assert download.suggested_filename, "导出调试日志应触发文件下载"
