import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


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


# 用例编号：TestCase_AcuHMI_003_05_case21
# 用例标题：post channel配置非法值或超过长度，系统提示字段验证错误
def test_TestCase_AcuHMI_003_05_case21(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_channel(page, 1)

    # Enable with FTP
    page.locator(".el-radio").filter(has_text="Enable").click()
    page.wait_for_timeout(300)

    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="FTP", exact=True).click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    # Test: invalid URL format (non-FTP URL)
    try:
        page.locator(".el-form-item").filter(has_text="FTP URL").locator("input").fill("999.999.999.999")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    has_url_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_url_error, "无效FTP URL格式应显示字段验证错误"

    # Test: port out of range
    try:
        page.locator(".el-form-item").filter(has_text="FTP URL").locator("input").fill("FTP://192.168.1.100")
    except Exception:
        pass
    try:
        port_input = page.locator(".el-form-item").filter(has_text="FTP Port").locator("input")
        port_input.fill("99999")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    has_port_error = (
        page.locator(".el-form-item__error").count() > 0
        or page.locator(".el-message--error").count() > 0
    )
    assert has_port_error, "端口超出范围（>65535）应显示字段验证错误"
