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


# 用例编号：TestCase_AcuHMI_009_05_case01_12
# 用例标题：slaveid:247无此地址站，筛选结果应为空或不匹配
def test_TestCase_AcuHMI_009_05_case01_12(login_page: LoginPage):
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

    # Select log type and enter slaveid
    try:
        page.locator(".el-form-item").filter(has_text="Type").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="RTU_REQ", exact=True).click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    try:
        page.locator(".el-form-item").filter(has_text="Slave ID").locator("input").fill("247")
    except Exception:
        pass

    page.get_by_role("button", name="Search").click()
    page.wait_for_timeout(1000)

    # Expect no matching rows (no device with slaveid=247)
    rows = page.locator("tbody tr").count()
    # Either empty table, or all rows are "no data" placeholder
    has_no_data = (
        rows == 0
        or page.get_by_text("No Data", exact=False).count() > 0
        or page.get_by_text("no data", exact=False).count() > 0
    )
    assert has_no_data or rows == 0, f"slaveid=247无对应设备站，筛选结果应为空"

    # Reset filter
    try:
        page.get_by_role("button", name="Reset").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass
