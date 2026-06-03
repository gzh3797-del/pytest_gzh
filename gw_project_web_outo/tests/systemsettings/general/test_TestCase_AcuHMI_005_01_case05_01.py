# 用例编号: TestCase_AcuHMI_005_01_case05_01
# 用例标题: 3个NTP服务器链接字符均为0个（空），保存配置失败
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. 添加3个NTP服务，服务器链接均为空，保存配置失败
# 预期结果: NTP服务器地址为空时保存失败，显示错误信息

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


def test_TestCase_AcuHMI_005_01_case05_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Date & Time")

    # 清空3个NTP服务器地址输入框
    page.get_by_placeholder("NTP Server 1").clear()
    page.get_by_placeholder("NTP Server 2").clear()
    page.get_by_placeholder("NTP Server 3").clear()
    page.get_by_role("button", name="Save").click()
    assert page.locator(".el-form-item__error").count() > 0, \
        "NTP服务器地址全为空时应保存失败并显示错误信息"
