# 用例编号: TestCase_AcuHMI_005_01_case05 (row 598)
# 用例标题: 3个NTP服务器链接字符分别超过40个，保存配置失败；
#           第1个NTP服务器链接为time.apple.co（有效但不可达），保存成功
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. 添加3个NTP服务，服务器链接均为41字符，保存配置失败
#   2. 添加第一个NTP服务，服务器链接为time.apple.co，保存配置成功，时间同步失败
# 预期结果:
#   步骤1: 保存失败，显示字段长度超限错误信息
#   步骤2: 保存成功，时间同步失败（因服务器不可达）

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


def test_TestCase_AcuHMI_005_01_case05_row598(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Date & Time")

    # 步骤1: 3个NTP服务器均填写41字符，保存应失败
    long_url = "qwertyuiopasdfghjklzxcvbnm123456789012345"  # 41字符
    assert len(long_url) == 41, f"long_url长度应为41，实际为{len(long_url)}"

    page.get_by_placeholder("NTP Server 1").fill(long_url)
    page.get_by_placeholder("NTP Server 2").fill(long_url)
    page.get_by_placeholder("NTP Server 3").fill(long_url)
    page.get_by_role("button", name="Save").click()
    assert page.locator(".el-form-item__error").count() > 0, \
        "3个NTP URL均超过40字符应保存失败"

    # 步骤2: 第1个NTP服务器填写time.apple.co（有效但不可达），其他清空，保存应成功
    page.get_by_placeholder("NTP Server 1").fill("time.apple.co")
    page.get_by_placeholder("NTP Server 2").fill("")
    page.get_by_placeholder("NTP Server 3").fill("")
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000), \
        "有效NTP服务器地址应保存成功（即使服务器不可达，保存本身应成功）"
