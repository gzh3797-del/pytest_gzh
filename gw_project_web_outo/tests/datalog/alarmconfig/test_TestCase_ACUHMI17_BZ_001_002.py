import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_ACUHMI17_BZ_001_002
# 用例标题：Alarm页面包含Active Alarms和Alarm Logs子页面
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 点击Alarm导航入口进入Alarm页面
#   2. 查看页面内子页面结构
# 预期结果：
#   2. 页面内正确展示Active Alarms和Alarm Logs两个子页面，均可正常跳转访问


def _nav_to_alarm(page, submenu: str = None):
    if "/#/alarm" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    if submenu:
        page.get_by_role("menuitem", name=submenu).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_ACUHMI17_BZ_001_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: Navigate to Alarm section
    _nav_to_alarm(page)
    page.wait_for_timeout(800)

    # Step 2: Verify the two sub-pages (Active Alarms and Alarm Logs) are shown

    # Check that the left-nav menu or tab area shows both sub-pages
    active_alarms_link = page.get_by_role("menuitem", name="Unacknowledged Alarms")
    alarm_logs_link = page.get_by_role("menuitem", name="Alarm Logs")

    expect(active_alarms_link).to_be_visible(), \
        "Alarm导航区域应显示'Unacknowledged Alarms'子页面入口"
    expect(alarm_logs_link).to_be_visible(), \
        "Alarm导航区域应显示'Alarm Logs'子页面入口"

    # Navigate to Active Alarms and verify URL
    active_alarms_link.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    assert "alarm" in page.url.lower() or "activeAlarm" in page.url, (
        f"点击'Unacknowledged Alarms'后URL应包含alarm相关路径，实际URL: {page.url}"
    )
    # Page should not show error
    assert page.locator(".el-message--error").count() == 0, \
        "Active Alarms页面不应出现错误提示"

    # Navigate to Alarm Logs and verify URL
    alarm_logs_link.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    assert "alarm" in page.url.lower() or "alarmLog" in page.url, (
        f"点击'Alarm Logs'后URL应包含alarm相关路径，实际URL: {page.url}"
    )
    # Page should not show error
    assert page.locator(".el-message--error").count() == 0, \
        "Alarm Logs页面不应出现错误提示"

    # Additionally verify that at least a table / data area renders in Alarm Logs
    # (the page should have a container element even if there are no alarms)
    page_body = page.locator(".el-table, .alarm-list, .data-table, main").first
    expect(page_body).to_be_visible()
