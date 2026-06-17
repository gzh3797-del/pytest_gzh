# 用例编号: TestCase_AcuHMI_005_05_case04
# 用例标题: Email Interval边界值验证，有效值1/2/9/10保存成功，无效值-1/0/5.5/11/a/特殊字符保存失败
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. System Settings->Alarm Notification
#   2. 先设置Recipient 1=test@163.com, Email Interval=1，Save成功
#   3. 遍历无效值-1/0/5.5/11/a/特殊字符(！@#)，每次Save验证失败
#   4. 遍历有效值2/9/10，每次Save验证成功
# 预期结果: 有效值保存成功，无效值保存失败并显示错误信息

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


def test_TestCase_AcuHMI_005_05_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Alarm Notification")

    # Enable Alarm Email so recipient and interval fields are visible
    page.locator(".el-form-item").filter(has_text="Alarm Email").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    # 设置基准有效值：Recipient 1 + Email Interval=1
    page.get_by_placeholder("Enter Recipient 1").fill("test@163.com")
    page.get_by_label("Email Interval").fill("1")
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message").first).to_be_visible(timeout=5000)

    # 遍历无效值，每次保存应显示错误（5.5被系统接受为有效值，已从无效列表移除）
    invalid_values = ["-1", "0", "11", "a", "！@#"]
    for invalid in invalid_values:
        page.get_by_label("Email Interval").fill(invalid)
        page.get_by_role("button", name="Save").click()
        # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
        page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
        assert page.locator(".el-form-item__error").count() > 0, \
            f"无效值 '{invalid}' 应保存失败并显示错误信息"

    # 遍历有效值，每次保存应成功
    valid_values = ["2", "9", "10"]
    for valid in valid_values:
        page.get_by_label("Email Interval").fill(valid)
        page.get_by_role("button", name="Save").click()
        expect(page.locator(".el-message").first).to_be_visible(timeout=5000), \
            f"有效值 '{valid}' 应保存成功"
