import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_physical_devices(page):
    if "/#/physicalDevice" not in page.url or "addDevice" in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_001_01_case12
# 用例标题：Download接入设备列表成功，设备列表数据准确
# 预置条件：
#   1. 接入设备支持Modbus RTU/TCP
# 测试步骤：
#   1. 在Physical Devices页面点击Download List按钮
# 预期结果：
#   1. 设备列表文件下载成功（CSV/Excel文件，文件名不为空）
def test_TestCase_AcuHMI_001_01_case12(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_physical_devices(page)

    # Trigger file download
    with page.expect_download() as dl_info:
        page.get_by_role("button", name="Download List").click()

    download = dl_info.value
    assert download.suggested_filename != "", \
        "Download List 应触发文件下载，且下载文件名不为空"

    # Optionally verify the file extension is CSV or Excel
    filename = download.suggested_filename.lower()
    assert any(filename.endswith(ext) for ext in (".csv", ".xlsx", ".xls")), \
        f"下载文件应为CSV或Excel格式，实际文件名: {download.suggested_filename}"
