import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

_DUPLICATE_SN = "SN123"
_FIRST_DEVICE_NAME = "UniqueTestDev01"
_SECOND_DEVICE_NAME = "UniqueTestDev02"


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


def _add_web_device(page, name: str, sn: str, model: str = "ModelX", url: str = "test.example.com") -> bool:
    """
    Open Add Device dialog, fill fields, click Confirm.
    Returns True if dialog closed (success), False if it stayed open (error).
    """
    page.get_by_role("button", name="Add Device").click()
    page.wait_for_timeout(500)

    dialog = page.locator(".el-dialog")
    dialog.wait_for(timeout=5000)

    def fill_by_label(text, value):
        inp = dialog.locator(".el-form-item").filter(has_text=text).locator("input").first
        inp.fill(value)

    fill_by_label("Device Name", name)
    fill_by_label("Serial Number", sn)
    fill_by_label("Model", model)
    dialog.locator("input[placeholder='---Enter URL---']").fill(url)
    page.wait_for_timeout(200)

    dialog.get_by_role("button", name="Confirm").click()
    page.wait_for_timeout(1500)

    # If dialog is still visible, the add failed
    return dialog.count() == 0 or not dialog.is_visible()


def _delete_device_by_name(page, device_name: str):
    """Delete a web device by its device name."""
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


# 用例编号：TestCase_AcuHMI_012_01_case09
# 用例标题：Serial Number唯一性校验，重复SN添加失败并提示
# 预置条件：1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 已存在Serial Number="SN123"
#   2. 打开Add Device，新建设备输入SN="SN123"
#   3. 点击Confirm
# 预期结果：
#   3. 无法保存，提示语正确
def test_TestCase_AcuHMI_012_01_case09(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_web_devices(page)

    try:
        # Step 1: Add the first device with SN="SN123"
        first_success = _add_web_device(
            page,
            name=_FIRST_DEVICE_NAME,
            sn=_DUPLICATE_SN,
            model="ModelA",
            url="device1.example.com",
        )
        assert first_success, \
            f"首次添加 SN={_DUPLICATE_SN} 的设备应成功，但弹窗未关闭"

        # Verify first device is visible in the list
        page.wait_for_timeout(500)
        expect(
            page.locator("tbody").get_by_role("row").filter(has_text=_FIRST_DEVICE_NAME)
        ).to_be_visible(timeout=5000)

        # Step 2: Attempt to add a second device with the same SN="SN123"
        page.get_by_role("button", name="Add Device").click()
        page.wait_for_timeout(500)

        dialog = page.locator(".el-dialog")
        dialog.wait_for(timeout=5000)
        expect(dialog).to_be_visible(timeout=5000)

        def fill_by_label2(text, value):
            inp = dialog.locator(".el-form-item").filter(has_text=text).locator("input").first
            inp.fill(value)

        fill_by_label2("Device Name", _SECOND_DEVICE_NAME)
        fill_by_label2("Serial Number", _DUPLICATE_SN)
        fill_by_label2("Model", "ModelB")
        dialog.locator("input[placeholder='---Enter URL---']").fill("device2.example.com")
        page.wait_for_timeout(200)

        # Step 3: Click Confirm — should fail due to duplicate SN
        dialog.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(1500)

        # Assert the operation was rejected
        dialog_still_visible = dialog.is_visible() if dialog.count() > 0 else False
        error_toast = page.locator(".el-message--error").count() > 0
        form_errors = dialog.locator(".el-form-item__error").count() if dialog.count() > 0 else 0

        assert dialog_still_visible or error_toast or form_errors > 0, \
            f"重复 SN={_DUPLICATE_SN} 的设备添加应失败并有提示，但未检测到任何错误信息"

        # If dialog is still open, assert there is an error message
        if dialog_still_visible:
            # There should be some error indication in the dialog or as a toast
            has_error = (
                dialog.locator(".el-form-item__error").count() > 0
                or page.locator(".el-message--error").count() > 0
                or dialog.get_by_text("already", exact=False).count() > 0
                or dialog.get_by_text("exist", exact=False).count() > 0
                or dialog.get_by_text("duplicate", exact=False).count() > 0
                or dialog.get_by_text("唯一", exact=False).count() > 0
                or dialog.get_by_text("重复", exact=False).count() > 0
            )
            assert has_error, \
                f"重复 SN={_DUPLICATE_SN} 的设备弹窗仍打开，但未发现错误提示语"

        # Dismiss dialog if still open
        if dialog.count() > 0 and dialog.is_visible():
            try:
                dialog.get_by_role("button", name="Cancel").click(timeout=2000)
                page.wait_for_timeout(500)
            except Exception:
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass

        # Verify the second device was NOT added (list should not contain second device name)
        second_device_row = page.locator("tbody tr").filter(has_text=_SECOND_DEVICE_NAME)
        assert second_device_row.count() == 0, \
            f"重复 SN 的设备 '{_SECOND_DEVICE_NAME}' 不应出现在设备列表中，但实际存在"

    finally:
        # Cleanup: delete the first device
        _delete_device_by_name(page, _FIRST_DEVICE_NAME)
        # Also clean up second device if it somehow got created
        _delete_device_by_name(page, _SECOND_DEVICE_NAME)
