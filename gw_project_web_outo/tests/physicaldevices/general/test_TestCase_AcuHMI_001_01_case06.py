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


def _get_field(page, label: str, placeholder: str):
    """Try label first, fall back to placeholder."""
    locator = page.get_by_label(label, exact=True)
    if locator.count() == 0:
        locator = page.get_by_placeholder(placeholder)
    return locator


def _assert_error(page, msg: str):
    has_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_error, msg


# 用例编号：TestCase_AcuHMI_001_01_case06
# 用例标题：添加接入设备时，"Serial Number"超过40个字符或者"Device Name"超过40个字符，无法保存
# 预置条件：
#   1. 接入设备支持Modbus RTU/TCP
# 测试步骤：
#   1. 配置Serial Number=41字符，保存
#   2. 配置Device Name=41字符，保存
#   3. 配置Serial Number含特殊字符@#$%，保存
#   4. 配置Device Name含特殊字符!@#$%，保存
# 预期结果：
#   2. 保存配置失败，错误提示Serial Number/Device Name长度超过40
#   4. 保存配置失败，错误提示不允许包含特殊字符
def test_TestCase_AcuHMI_001_01_case06(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # --- Step 1: Serial Number = 41 characters ---
    _open_add_device_form(page)
    serial_field = _get_field(page, "Serial Number", "Enter Serial Number")
    serial_field.clear()
    serial_field.fill("S" * 41)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_error(page, "Serial Number=41字符时应显示验证错误（超过最大长度40），但未检测到错误提示")

    # --- Step 2: Device Name = 41 characters ---
    _open_add_device_form(page)
    name_field = _get_field(page, "Device Name", "Enter Device Name")
    name_field.clear()
    name_field.fill("D" * 41)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_error(page, "Device Name=41字符时应显示验证错误（超过最大长度40），但未检测到错误提示")

    # --- Step 3: Serial Number with special characters @#$% ---
    _open_add_device_form(page)
    serial_field2 = _get_field(page, "Serial Number", "Enter Serial Number")
    serial_field2.clear()
    serial_field2.fill("SN@#$%")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_error(page, "Serial Number含特殊字符@#$%时应显示验证错误，但未检测到错误提示")

    # --- Step 4: Device Name with special characters !@#$% ---
    _open_add_device_form(page)
    name_field2 = _get_field(page, "Device Name", "Enter Device Name")
    name_field2.clear()
    name_field2.fill("Dev!@#$%")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    _assert_error(page, "Device Name含特殊字符!@#$%时应显示验证错误，但未检测到错误提示")
