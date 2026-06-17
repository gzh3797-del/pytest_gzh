# 用例编号: TestCase_AcuHMI_005_04_case06_3
# 用例标题: 密码为41个字符，保存失败
# 预置条件: 管理权限登录AcuHMI网页
# 测试步骤:
#   1. 进入 System Settings -> Email
#   2. 填写基准配置，将 Password 改为41个字符（超出最大40字符限制）
#   3. 点击 Save
# 预期结果: 保存失败，显示表单校验错误

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


def test_TestCase_AcuHMI_005_04_case06_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Email")
    _fill_email_baseline(page)
    page.get_by_label("Password", exact=True).fill("q" * 41)
    page.get_by_role("button", name="Save").click()
    # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
    page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
    assert page.locator(".el-form-item__error").count() > 0
