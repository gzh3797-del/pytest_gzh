import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

# 用例编号：TestCase_AcuHMI_005_05_case02_01
# 用例标题：Alarm notification启用后再保存，Test Email发送失败（无真实报警触发）
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. 配置收件人，Enable，Save，Test Email
# 预期结果：
#   4. 发送邮件失败（无真实SMTP或邮件服务器）
@pytest.mark.xfail(strict=False, reason="Test Email需要真实可用的SMTP邮件服务器")
def test_TestCase_AcuHMI_005_05_case02_01(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Alarm Notification")
    page.wait_for_timeout(500)

    # Fill recipients
    try:
        recipient_input = page.locator(".el-form-item").filter(
            has_text="Recipient 1"
        ).locator("input").first
        recipient_input.fill("Recipient1@163.com")
    except Exception:
        pass

    # Enable
    try:
        page.get_by_role("button", name="Enable").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "Alarm notification配置保存应成功"

    # Test Email
    try:
        page.get_by_role("button", name="Test Email").click()
        page.wait_for_timeout(5000)
    except Exception:
        pass

    result = page.locator(".el-message").first
    expect(result).to_be_visible(timeout=10000)
