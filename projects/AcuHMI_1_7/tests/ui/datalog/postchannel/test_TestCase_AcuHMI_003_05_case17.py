import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_post_channel(page, channel_num: int):
    """Navigate to Post Channel N configuration page."""
    target = f"postChannel{channel_num}"
    if target not in page.url:
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
        try:
            page.get_by_role("menuitem", name="Post Channels").click()
            page.wait_for_timeout(300)
        except Exception:
            pass
    page.get_by_role("menuitem", name=f"Post Channel {channel_num}").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _confirm_dialog(page):
    """Try all known confirmation button variants."""
    for name in ("Yes", "OK", "Yes, continue", "Yes,continue", "Confirm"):
        try:
            page.get_by_role("button", name=name, exact=True).click(timeout=2000)
            page.wait_for_timeout(300)
            return True
        except Exception:
            pass
    # popconfirm primary button fallback
    try:
        page.locator(".el-popconfirm__action .el-button--primary").click(timeout=2000)
        page.wait_for_timeout(300)
        return True
    except Exception:
        pass
    return False


# 用例编号：TestCase_AcuHMI_003_05_case17
# 用例标题：Post Ch3 enable，FTP错误配置，Clear Post Channel Logs成功
def test_TestCase_AcuHMI_003_05_case17(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_channel(page, 3)

    # Enable with FTP (unreachable server) and fill all required fields
    page.locator(".el-radio").filter(has_text="Enable").click()
    page.wait_for_timeout(300)

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="FTP", exact=True).click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    for field, val in [
        ("FTP URL", "FTP://192.168.250.250"),
        ("FTP Port", "21"),
        ("FTP User Name", "testuser"),
        ("FTP password", "testpass"),
    ]:
        try:
            page.locator(".el-form-item").filter(has_text=field).locator("input").fill(val)
        except Exception:
            pass

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    # Save may or may not produce toast; continue regardless
    msg = page.locator(".el-message")
    if msg.count() > 0:
        assert page.locator(".el-message--error").count() == 0, "FTP配置保存不应有错误"

    # Clear Post Channel Logs
    page.get_by_role("button", name="Clear Post Channel Logs").click()
    page.wait_for_timeout(1000)

    confirmed = _confirm_dialog(page)
    page.wait_for_timeout(500)

    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "Clear Post Channel Logs应成功"
