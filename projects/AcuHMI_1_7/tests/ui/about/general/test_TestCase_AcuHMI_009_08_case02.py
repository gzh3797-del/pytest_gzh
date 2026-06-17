import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case02
# 用例标题：设备位置大于40个字符，保存配置失败，小于等于40个字符，保存配置成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. About->Device information设置Location超过40个字符，点击Save
#   2. 设置Location等于40个字符，点击Save
# 预期结果：
#   1. 弹框提示"Location cannot exceed 40 characters"
#   2. 弹框提示"Device info saved"
def test_TestCase_AcuHMI_009_08_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # Step 1: 输入超过40个字符的Location，点击Save，验证弹框提示"Location cannot exceed 40 characters"
    location_field = page.get_by_placeholder("Enter Location")
    location_field.clear()
    location_field.fill("b" * 41)
    page.get_by_role("button", name="Save").click()
    expect(page.get_by_text("Location cannot exceed 40 characters", exact=False)).to_be_visible(timeout=3000)

    # Step 2: 输入恰好等于40个字符的Location，点击Save，验证弹框提示"Device info saved"
    location_field.clear()
    location_field.fill("b" * 40)
    page.get_by_role("button", name="Save").click()
    expect(page.get_by_text("Device info saved", exact=False)).to_be_visible(timeout=5000)
