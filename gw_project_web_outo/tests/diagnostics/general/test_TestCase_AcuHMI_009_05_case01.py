import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_diagnostics(page, submenu: str):
    if "/diagnostics" not in page.url:
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


# 用例编号：TestCase_AcuHMI_009_05_case01
# 用例标题：禁用Modbus调试跟踪，Modbus调试不跟踪，导出调试日志导出成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Diagnostics->Modbus Debug Log设置Modbus Debug Trace为Disable
#   2. 点击Export Debug Logs
# 预期结果：
#   2. 调试日志导出成功
def test_TestCase_AcuHMI_009_05_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Modbus Debug Log")

    # Disable Modbus Debug Trace (if currently enabled)
    try:
        page.get_by_role("button", name="Disable").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    with page.expect_download(timeout=15000) as download_info:
        page.get_by_role("button", name="Export Debug Logs").click()
    download = download_info.value
    assert download.suggested_filename, "导出调试日志应触发文件下载，文件名不为空"
