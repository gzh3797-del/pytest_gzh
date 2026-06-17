# 用例编号: TestCase_AcuHMI_005_05_case05_3
# 用例标题: 三个收件人邮件长度均为41字符，保存配置成功
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. System Settings->Alarm Notification
#   2. 三个Recipient均为0123456789012345678901234567890123@qq.com (41字符)，
#      Email Interval=5, Save
# 预期结果: 41字符有效邮件地址配置保存成功

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


def test_TestCase_AcuHMI_005_05_case05_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Alarm Notification")

    # Enable Alarm Email so recipient fields appear
    page.locator(".el-form-item").filter(has_text="Alarm Email Enable").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    # 构造41字符邮件地址: "0123456789012345678901234567890123@qq.com"
    email_41 = "0123456789012345678901234567890123@qq.com"
    assert len(email_41) == 41, f"email_41长度应为41，实际为{len(email_41)}"

    # 三个Recipient均填写41字符邮件地址，Email Interval=5，保存
    page.get_by_placeholder("Enter Recipient 1").fill(email_41)
    page.get_by_placeholder("Enter Recipient 2").fill(email_41)
    page.get_by_placeholder("Enter Recipient 3").fill(email_41)
    page.get_by_label("Email Interval").fill("5")
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000), \
        "41字符有效邮件地址应配置保存成功"
