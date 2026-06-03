# 用例编号: TestCase_AcuHMI_005_05_case05_2
# 用例标题: 三个收件人邮件长度均为5字符（太短），保存配置失败
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. System Settings->Alarm Notification
#   2. Recipient1=@1com, Recipient2=qq.com, Recipient3=3@.com,
#      Email Interval=5, Save
# 预期结果: 过短邮件地址保存失败，显示错误信息

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


def test_TestCase_AcuHMI_005_05_case05_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Alarm Notification")

    # Enable Alarm Email so recipient fields appear
    page.locator(".el-form-item").filter(has_text="Alarm Email Enable").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    # 填写长度为5字符的过短邮件地址
    page.get_by_placeholder("Enter Recipient 1").fill("@1com")
    page.get_by_placeholder("Enter Recipient 2").fill("qq.com")
    page.get_by_placeholder("Enter Recipient 3").fill("3@.com")
    page.get_by_label("Email Interval").fill("5")
    page.get_by_role("button", name="Save").click()
    assert page.locator(".el-form-item__error").count() > 0, \
        "过短邮件地址应保存失败并显示错误信息"
