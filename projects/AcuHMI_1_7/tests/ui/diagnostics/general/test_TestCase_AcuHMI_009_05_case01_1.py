import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


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


# 用例编号：TestCase_AcuHMI_009_05_case01_1
# 用例标题：禁用Modbus调试跟踪，Modbus调试不跟踪，清除调试日志清除成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Diagnostics->Modbus Debug Log设置Modbus Debug Trace为Disable
#   2. 点击Clear Debug Logs
# 预期结果：
#   2. 调试日志清除成功
def test_TestCase_AcuHMI_009_05_case01_1(login_page: LoginPage):
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

    page.get_by_role("button", name="Clear Debug Logs").click()
    page.wait_for_timeout(1000)

    # Handle confirmation dialog — button text has no space: "Yes,continue"
    try:
        page.get_by_role("button", name="Yes,continue").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.get_by_role("button", name="Confirm").click(timeout=2000)
            page.wait_for_timeout(500)
        except Exception:
            pass

    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
