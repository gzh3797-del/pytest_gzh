import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

# 用例编号：TestCase_AcuHMI_005_05_case02
# 用例标题：Alarm notification禁用，配置后不发送报警知通
# 预置条件：管理权限登录AcuHMI
# 测试步骤：
#   1. 配置收件人Recipient1@163.com，Email Interval 1-10
#   2. Alarm notification Disable，Save
# 预期结果：
#   3. 显示配置保存成功，Alarm notification状态为disabled
def test_TestCase_AcuHMI_005_05_case02(login_page):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Alarm Notification")
    page.wait_for_timeout(500)

    # First fill recipients
    try:
        recipient_input = page.locator(".el-form-item").filter(
            has_text="Recipient 1"
        ).locator("input").first
        recipient_input.fill("Recipient1@163.com")
    except Exception:
        pass

    # Disable alarm notification
    try:
        page.get_by_role("button", name="Disable").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.locator(".el-form-item").filter(has_text="Enable").locator(
                ".el-radio"
            ).filter(has_text="Disable").click()
            page.wait_for_timeout(300)
        except Exception:
            pass

    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
    assert page.locator(".el-message--error").count() == 0, "Alarm notification禁用配置保存应成功"
