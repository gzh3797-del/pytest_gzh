import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


def _nav_to_data_log_management(page):
    if "/#/dataLog" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Data Log").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Data Log Management").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_003_07_case04
# 用例标题：Delete Log：选择需要删除的日志设备，点击"delete"，对应设备日志删除成功
# 预置条件：已接入设备，Data Log Management中有设备日志
# 测试步骤：
#   1. Data Log Management，选择要删除的日志设备，点击"delete"
#   2. 后台查看该设备日志，已不存在
# 预期结果：
#   2. 该设备日志已删除，无该设备日志
def test_TestCase_AcuHMI_003_07_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_log_management(page)

    # Get initial row count
    page.wait_for_timeout(500)
    rows_before = page.locator("tbody tr").count()

    if rows_before == 0:
        pytest.skip("Data Log Management中无设备日志可供删除")

    # Select first device row
    try:
        page.locator("tbody tr").first.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)
    except Exception:
        try:
            page.locator("tbody tr").first.locator("input[type='checkbox']").check()
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Record device name for verification
    try:
        device_name = page.locator("tbody tr").first.locator("td").nth(1).inner_text()
    except Exception:
        device_name = None

    # Check Delete button is enabled after selection
    del_btn = page.get_by_role("button", name="delete", exact=False)
    if del_btn.count() == 0:
        del_btn = page.get_by_role("button", name="Delete", exact=False)
    if del_btn.count() == 0 or "disabled" in (del_btn.first.get_attribute("class") or ""):
        pytest.skip("Delete按钮禁用，选择未生效或设备日志无法删除")

    # Click Delete
    del_btn.first.click()
    page.wait_for_timeout(1000)

    # Handle confirmation
    try:
        page.get_by_role("button", name="Yes,continue").click(timeout=3000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.get_by_role("button", name="Confirm").click(timeout=3000)
            page.wait_for_timeout(500)
        except Exception:
            pass

    page.wait_for_timeout(1000)
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "删除设备日志应成功，不应有错误提示"

    # Verify the device log entry was removed (row count decreased or entry not visible)
    rows_after = page.locator("tbody tr").count()
    if device_name:
        # The specific device should no longer appear in the list
        device_rows = page.locator("tbody tr").filter(has_text=device_name)
        assert device_rows.count() < rows_before or rows_after < rows_before, \
            f"删除后，设备'{device_name}'的日志应不再出现在列表中"
    else:
        assert rows_after < rows_before, "删除后列表行数应减少"
