# 用例编号: TestCase_AcuHMI_005_04_case04_5
# 用例标题: Email Server=aaa.aa.a.a，Port=#，保存失败
# 预置条件: 管理权限登录AcuHMI网页
# 测试步骤:
#   1. 进入 System Settings -> Email
#   2. 配置 Email Server=aaa.aa.a.a, Email Port=#（特殊字符）
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


def test_TestCase_AcuHMI_005_04_case04_5(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Email")
    page.get_by_label("Email Server", exact=False).fill("aaa.aa.a.a")
    page.get_by_label("Email Port", exact=False).fill("#")
    page.get_by_role("button", name="Save").click()
    # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
    page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
    assert page.locator(".el-form-item__error").count() > 0
