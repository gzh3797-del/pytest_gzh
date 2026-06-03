import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case01
# 用例标题：设备名字大于40个字符，保存配置失败，小于等于40个字符，保存配置成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. About->Device information设置name超过40个字符，点击Save
#   2. 设置name等于40个字符，点击Save
# 预期结果：
#   1. 弹框提示"Name cannot exceed 40 characters"
#   2. 弹框提示"Device info saved"
def test_TestCase_AcuHMI_009_08_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # Step 1: 输入超过40个字符的Name，点击Save，验证弹框提示"Name cannot exceed 40 characters"
    name_field = page.get_by_placeholder("Enter Name")
    name_field.clear()
    name_field.fill("a" * 41)
    page.get_by_role("button", name="Save").click()
    expect(page.get_by_text("Name cannot exceed 40 characters", exact=False)).to_be_visible(timeout=3000)

    # Step 2: 输入恰好等于40个字符的Name，点击Save，验证弹框提示"Device info saved"
    name_field.clear()
    name_field.fill("a" * 40)
    page.get_by_role("button", name="Save").click()
    expect(page.get_by_text("Device info saved", exact=False)).to_be_visible(timeout=5000)
