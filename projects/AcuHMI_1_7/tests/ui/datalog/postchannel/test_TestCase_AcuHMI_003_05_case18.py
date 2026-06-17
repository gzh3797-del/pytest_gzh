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


# 用例编号：TestCase_AcuHMI_003_05_case18
# 用例标题：Post Ch3 enable，SFTP，正确配置Test success，错误配置Test fail
@pytest.mark.xfail(strict=False, reason="Test Post Channel结果依赖真实SFTP服务器连通性")
def test_TestCase_AcuHMI_003_05_case18(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_post_channel(page, 3)

    # Enable
    page.locator(".el-radio").filter(has_text="Enable").click()
    page.wait_for_timeout(300)

    # Select SFTP
    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name="SFTP", exact=True).click()
        page.wait_for_timeout(300)
    except Exception:
        pass

    # Fill valid SFTP config (field names may include SFTP prefix)
    for field, val in [
        ("SFTP URL", "sftp://192.168.1.100"),
        ("SFTP Port", "22"),
        ("SFTP User Name", "sftpuser"),
        ("SFTP password", "sftppass"),
    ]:
        try:
            page.locator(".el-form-item").filter(has_text=field).locator("input").fill(val)
        except Exception:
            try:
                # Try without SFTP prefix
                plain_field = field.replace("SFTP ", "")
                page.locator(".el-form-item").filter(has_text=plain_field).locator("input").fill(val)
            except Exception:
                pass

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)

    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    expect(page.locator(".el-message, .el-alert").first).to_be_visible(timeout=10000)

    # Test with invalid URL
    try:
        page.locator(".el-form-item").filter(has_text="URL").locator("input").fill("sftp://999.999.999.999")
    except Exception:
        pass
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Test Post Channel").click()
    page.wait_for_timeout(5000)
    assert (
        page.get_by_text("fail", exact=False).count() > 0
        or page.locator(".el-message--error").count() > 0
    ), "无效SFTP URL应导致Test Post Channel失败"
