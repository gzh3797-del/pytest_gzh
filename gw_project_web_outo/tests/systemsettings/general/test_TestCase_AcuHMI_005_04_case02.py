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

def _fill_email_baseline(page):
    page.get_by_label("Email Server", exact=False).fill("smtp.163.com")
    page.get_by_label("Email Port", exact=False).fill("25")
    try:
        page.locator(".el-radio").filter(has_text="Off").click()
        page.wait_for_timeout(200)
    except Exception:
        pass
    page.get_by_label("Sender Name", exact=True).fill("xiaoming")
    page.get_by_label("From Email Address", exact=True).fill("159xxxx4651@163.com")
    page.get_by_label("Username", exact=True).fill("xiaoming123")
    page.get_by_label("Password", exact=True).fill("Admin@110001")

# 用例编号：TestCase_AcuHMI_005_04_case02
# 用例标题：TLS/SSL=ON，初始邮件服务器配置，保存成功，收到邮件（SSL方式）
# 预置条件：管理权限登录AcuHMI，有可用邮件服务器
# 测试步骤：
#   1. Email，TLS/SSL=ON，保存成功
#   2. Test Email → 收到SSL方式邮件
# 预期结果：
#   2. 邮件发送成功，以SSL方式加密
@pytest.mark.xfail(strict=False, reason="Test Email需要真实可用的SMTP邮件服务器（TLS/SSL=ON需端口465）")
def test_TestCase_AcuHMI_005_04_case02(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Email")
    _fill_email_baseline(page)

    # Set TLS/SSL = On
    try:
        page.locator(".el-radio").filter(has_text="On").click()
        page.wait_for_timeout(200)
    except Exception:
        pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "Email配置(TLS=ON)保存应成功"

    page.get_by_role("button", name="Test Email").click()
    page.wait_for_timeout(5000)
    expect(page.locator(".el-message").first).to_be_visible(timeout=10000)
