import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


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


def _open_add_device_form(page):
    _nav_to_physical_devices(page)
    page.get_by_role("button", name="Add Device").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _fill_rtu_base_fields(page, device_name: str = "TestRTU_timeout"):
    """Fill the minimum required fields for an RTU device."""
    try:
        page.get_by_label("Device Name", exact=True).fill(device_name)
    except Exception:
        page.get_by_placeholder("Enter Device Name").fill(device_name)
    # Select RTU protocol if there is a protocol selector
    try:
        page.get_by_label("Protocol", exact=True).click()
        page.wait_for_timeout(300)
        page.get_by_role("option", name="RTU").click()
        page.wait_for_timeout(300)
    except Exception:
        pass


# 用例编号：TestCase_AcuHMI_001_01_case05
# 用例标题：添加接入设备时，"Request Timeout"设置0.001或5.1，无法保存
# 预置条件：
#   1. 接入设备支持Modbus RTU/TCP
#   2. 接入设备与AcuHMI物理连线正常
# 测试步骤：
#   1. 通过HMI Web页面RTU方式添加设备，Request Timeout=0.001
#   2. 检查保存是否成功
#   3. RTU方式添加，Request Timeout=5.1
#   4. 检查保存是否成功
# 预期结果：
#   2. 配置信息保存失败（0.001低于最小值0.1）
#   4. 配置信息保存失败（5.1超过最大值5.0）
def test_TestCase_AcuHMI_001_01_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # --- Step 1 & 2: Request Timeout = 0.001 (below minimum 0.1) ---
    _open_add_device_form(page)
    _fill_rtu_base_fields(page, "TestRTU_timeout_low")

    timeout_field = page.get_by_label("Request Timeout", exact=True)
    if timeout_field.count() == 0:
        timeout_field = page.get_by_placeholder("Enter Request Timeout")
    timeout_field.clear()
    timeout_field.fill("0.001")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)

    has_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_error, \
        "Request Timeout=0.001 (低于最小值) 时应显示验证错误，但未检测到错误提示"

    # --- Step 3 & 4: Request Timeout = 5.1 (above maximum 5.0) ---
    # Navigate back to add device form for a clean state
    _open_add_device_form(page)
    _fill_rtu_base_fields(page, "TestRTU_timeout_high")

    timeout_field2 = page.get_by_label("Request Timeout", exact=True)
    if timeout_field2.count() == 0:
        timeout_field2 = page.get_by_placeholder("Enter Request Timeout")
    timeout_field2.clear()
    timeout_field2.fill("5.1")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)

    has_error2 = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_error2, \
        "Request Timeout=5.1 (超过最大值) 时应显示验证错误，但未检测到错误提示"
