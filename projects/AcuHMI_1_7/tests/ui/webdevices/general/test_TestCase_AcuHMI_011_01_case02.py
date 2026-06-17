import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_DEVICE_NAME = "A"
_DEVICE_SN = "S1"
_DEVICE_MODEL = "M1"
_DEVICE_URL = "abc.com"


def _nav_to_web_devices(page):
    """Navigate to Devices > Web Devices."""
    if "/#/webDevice" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Web Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _delete_device_by_name(page, device_name: str):
    """Delete a web device by its name from the list."""
    _nav_to_web_devices(page)
    page.wait_for_timeout(500)
    row = page.locator("tbody tr").filter(has_text=device_name)
    if row.count() == 0:
        return
    try:
        row.get_by_role("button").last.click()
        page.wait_for_timeout(500)
        try:
            page.get_by_role("button", name="Yes").click(timeout=2000)
            page.wait_for_timeout(800)
        except Exception:
            pass
    except Exception:
        pass


# 用例编号：TestCase_AcuHMI_011_01_case02
# 用例标题：成功创建设备（最小输入长度）
# 预置条件：1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 打开Add Device弹窗
#   2. 输入Device Name: "A", SN: "S1", Model: "M1", URL: "https://abc.com"
#   3. 点击Confirm
# 预期结果：
#   3. 成功创建设备，弹窗关闭
def test_TestCase_AcuHMI_011_01_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_web_devices(page)

    try:
        # Open Add Device dialog
        page.get_by_role("button", name="Add Device").click()
        page.wait_for_timeout(500)

        dialog = page.locator(".el-dialog")
        dialog.wait_for(timeout=5000)
        expect(dialog).to_be_visible(timeout=5000)

        # 按 form-item label 文本定位各输入框（与 add_max_limit 相同方式）
        def fill_by_label(text, value):
            inp = dialog.locator(".el-form-item").filter(has_text=text).locator("input").first
            inp.fill(value)

        fill_by_label("Device Name", _DEVICE_NAME)
        fill_by_label("Serial Number", _DEVICE_SN)
        fill_by_label("Model", _DEVICE_MODEL)
        dialog.locator("input[placeholder='---Enter URL---']").fill(_DEVICE_URL)
        page.wait_for_timeout(200)

        # Click Confirm
        dialog.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(1500)

        # Assert dialog closed (success)
        dialog_visible = dialog.is_visible() if dialog.count() > 0 else False
        assert not dialog_visible, \
            "Add Device 弹窗应在 Confirm 后关闭，但弹窗仍可见（创建可能失败）"

        # Assert no error toast
        assert page.locator(".el-message--error").count() == 0, \
            "创建设备后不应出现错误 toast"

        # Assert the new device appears in the list
        page.wait_for_timeout(500)
        expect(page.locator("tbody").get_by_role("row").filter(has_text=_DEVICE_NAME)).to_be_visible(
            timeout=5000
        )

    finally:
        _delete_device_by_name(page, _DEVICE_NAME)
