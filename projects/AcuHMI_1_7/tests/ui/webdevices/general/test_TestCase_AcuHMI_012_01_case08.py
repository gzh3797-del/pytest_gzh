import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


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


def _open_add_dialog(page):
    """Open the Add Device dialog and return the dialog locator."""
    page.get_by_role("button", name="Add Device").click()
    page.wait_for_timeout(500)
    dialog = page.locator(".el-dialog")
    dialog.wait_for(timeout=5000)
    expect(dialog).to_be_visible(timeout=5000)
    return dialog


def _fill_and_confirm(page, dialog, name: str, sn: str, model: str, url: str):
    """Fill dialog fields and click Confirm."""
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


def _is_rejected(page, dialog) -> bool:
    """Return True if the operation was rejected (dialog still open or error shown)."""
    # Form errors in the dialog
    if dialog.count() > 0 and dialog.is_visible():
        form_errors = dialog.locator(".el-form-item__error").count()
        if form_errors > 0:
            return True
        # Dialog still open without errors — treat as rejected too (server-side check)
        return True
    # Error toast
    if page.locator(".el-message--error").count() > 0:
        return True
    return False


def _dismiss_dialog(page, dialog):
    """Close the dialog if still open."""
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


# 用例编号：TestCase_AcuHMI_012_01_case08
# 用例标题：长度校验和特殊字符校验（超长/特殊字符添加失败）
# 预置条件：1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 打开Add Device弹窗
#   2. 输入Device Name>40字符，SN>40字符，Model>40字符，URL>300字符
#   3. 点击Confirm
#   4. 输入特殊字符@#$%
# 预期结果：
#   3. 创建失败，提示语正确
#   4. 创建失败，提示语正确
def test_TestCase_AcuHMI_012_01_case08(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_web_devices(page)

    # --- Sub-test 1: Overly long Device Name (41 chars > max 40) ---
    dialog = _open_add_dialog(page)
    _fill_and_confirm(
        page, dialog,
        name="D" * 41,          # 41 chars — over the 40-char limit
        sn="ValidSN01",
        model="ValidModel01",
        url="valid.example.com",
    )
    assert _is_rejected(page, dialog), \
        "Device Name=41字符（超出40字符上限），保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)

    # --- Sub-test 2: Overly long Serial Number (41 chars > max 40) ---
    _nav_to_web_devices(page)
    dialog = _open_add_dialog(page)
    _fill_and_confirm(
        page, dialog,
        name="ValidName02",
        sn="S" * 41,             # 41 chars — over limit
        model="ValidModel02",
        url="valid.example.com",
    )
    assert _is_rejected(page, dialog), \
        "Serial Number=41字符（超出40字符上限），保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)

    # --- Sub-test 3: Overly long Model (41 chars > max 40) ---
    _nav_to_web_devices(page)
    dialog = _open_add_dialog(page)
    _fill_and_confirm(
        page, dialog,
        name="ValidName03",
        sn="ValidSN03",
        model="M" * 41,          # 41 chars — over limit
        url="valid.example.com",
    )
    assert _is_rejected(page, dialog), \
        "Model=41字符（超出40字符上限），保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)

    # --- Sub-test 4: Overly long URL (293 chars > max 292 input / 300 full URL) ---
    _nav_to_web_devices(page)
    dialog = _open_add_dialog(page)
    long_url = "ab." * 96 + "abcdef"   # 288+6=294 chars > 292 input limit
    _fill_and_confirm(
        page, dialog,
        name="ValidName04",
        sn="ValidSN04",
        model="ValidModel04",
        url=long_url,
    )
    assert _is_rejected(page, dialog), \
        f"URL输入={len(long_url)}字符（超出292字符上限），保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)

    # --- Sub-test 5: Special characters in Device Name ---
    _nav_to_web_devices(page)
    dialog = _open_add_dialog(page)
    _fill_and_confirm(
        page, dialog,
        name="Name@#$%",         # special characters
        sn="ValidSN05",
        model="ValidModel05",
        url="valid.example.com",
    )
    assert _is_rejected(page, dialog), \
        "Device Name含特殊字符@#$%，保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)

    # --- Sub-test 6: Special characters in Serial Number ---
    _nav_to_web_devices(page)
    dialog = _open_add_dialog(page)
    _fill_and_confirm(
        page, dialog,
        name="ValidName06",
        sn="SN@#$%",             # special characters
        model="ValidModel06",
        url="valid.example.com",
    )
    assert _is_rejected(page, dialog), \
        "Serial Number含特殊字符@#$%，保存应失败，但弹窗已关闭"
    _dismiss_dialog(page, dialog)
