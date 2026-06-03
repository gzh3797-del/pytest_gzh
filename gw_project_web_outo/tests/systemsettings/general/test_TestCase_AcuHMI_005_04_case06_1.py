# 用例编号: TestCase_AcuHMI_005_04_case06_1
# 用例标题: 密码为混合字符，隐藏状态查看，保存成功
# 预置条件: 管理权限登录AcuHMI网页
# 测试步骤:
#   1. 进入 System Settings -> Email
#   2. 填写基准配置，将 Password 改为 !@#AbC123
#   3. 不点击显示按钮，密码保持隐藏状态
#   4. 验证 input type 仍为 password（密码隐藏）
#   5. 点击 Save
# 预期结果: 密码保持隐藏状态，保存成功

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


def test_TestCase_AcuHMI_005_04_case06_1(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Email")
    _fill_email_baseline(page)
    pwd_value = "!@#AbC123"
    page.get_by_label("Password", exact=True).fill(pwd_value)
    # Do not click show toggle — password remains hidden
    pwd_input = page.get_by_label("Password", exact=True)
    input_type = pwd_input.get_attribute("type")
    assert input_type == "password", f"密码应处于隐藏状态，当前type：{input_type}"
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
