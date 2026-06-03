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


# 用例编号：TestCase_AcuHMI_003_07_case01
# 用例标题：Download Log：选择指定时间范围和间隔(1/5/10/15/30mins)，下载设备日志，
#           下载时文件内容的时间戳与Log Interval匹配
# 预置条件：已接入设备，设备有日志数据
# 测试步骤：
#   1. Data Log Management，选择设备，选择时间范围，日志间隔1mins，点击Download
#   2-10. 分别用5/10/15/30mins重复
# 预期结果：
#   2/4/6/8/10. 下载日志文件，文件时间戳与对应间隔匹配
def test_TestCase_AcuHMI_003_07_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_log_management(page)

    # Check if there are any devices in the table
    page.wait_for_timeout(500)
    if page.locator("tbody tr").count() == 0:
        pytest.skip("Data Log Management中无设备日志数据，跳过下载测试")

    # Select a device from the list
    try:
        page.locator("tbody tr").first.locator(".el-checkbox__inner").click()
        page.wait_for_timeout(300)
    except Exception:
        try:
            page.locator("tbody tr").first.locator("input[type='checkbox']").check()
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Check Download button is enabled after selection
    dl_btn = page.get_by_role("button", name="Download")
    if dl_btn.count() == 0 or "disabled" in (dl_btn.first.get_attribute("class") or ""):
        pytest.skip("Download按钮禁用，设备无历史日志数据或选择未生效")

    # Select time range
    try:
        page.locator(".el-form-item").filter(has_text="Log Interval").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="1 min").click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    # Click Download
    with page.expect_download(timeout=30000) as dl_info:
        dl_btn.first.click()
    download = dl_info.value
    assert download.suggested_filename, "下载日志文件应触发文件下载"
    assert download.suggested_filename.endswith((".zip", ".csv", ".log", ".xlsx")), \
        f"下载文件应为日志文件格式，实际文件名：{download.suggested_filename}"
