import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# Maximum boundary values per field specification
_DEVICE_NAME = "D" * 40          # 40 characters
_DEVICE_SN = "S" * 40            # 40 characters
_DEVICE_MODEL = "M" * 40         # 40 characters
# URL 输入框只填域名部分（表单自带 https:// 前缀共8字符），最大292字符
# https://(8) + 292 = 300字符，对应表单"Maximum 300 characters"
_DEVICE_URL = "ab." * 96 + "abcd"    # 288 + 4 = 292 chars


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


def _delete_device_by_sn(page, sn: str):
    """Delete a web device by its serial number."""
    _nav_to_web_devices(page)
    page.wait_for_timeout(500)
    row = page.locator("tbody tr").filter(has_text=sn)
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


# 用例编号：TestCase_AcuHMI_011_01_case03
# 用例标题：成功创建设备（最大边界值）
# 预置条件：1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 打开Add Device弹窗
#   2. 输入Device Name=40字符，SN=40字符，Model=40字符，URL=300字符
#   3. 点击Confirm
# 预期结果：
#   3. 创建成功，无报错
def test_TestCase_AcuHMI_011_01_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Verify our boundary strings are the correct length
    assert len(_DEVICE_NAME) == 40, f"Device Name 应为40字符，实际: {len(_DEVICE_NAME)}"
    assert len(_DEVICE_SN) == 40, f"Serial Number 应为40字符，实际: {len(_DEVICE_SN)}"
    assert len(_DEVICE_MODEL) == 40, f"Model 应为40字符，实际: {len(_DEVICE_MODEL)}"
    assert len(_DEVICE_URL) == 292, f"URL 输入部分应为292字符（+https://=300），实际: {len(_DEVICE_URL)}"

    _nav_to_web_devices(page)

    try:
        # Open Add Device dialog
        page.get_by_role("button", name="Add Device").click()
        page.wait_for_timeout(500)

        dialog = page.locator(".el-dialog")
        dialog.wait_for(timeout=5000)
        expect(dialog).to_be_visible(timeout=5000)

        # 按 form-item label 文本定位各输入框
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
            "Add Device 弹窗应在 Confirm 后关闭（最大边界值创建应成功），但弹窗仍可见"

        # Assert no form errors
        assert page.locator(".el-form-item__error").count() == 0, \
            "最大边界值创建设备后不应出现表单校验错误"

        # Assert no error toast
        assert page.locator(".el-message--error").count() == 0, \
            "最大边界值创建设备后不应出现错误 toast"

    finally:
        _delete_device_by_sn(page, _DEVICE_SN)
