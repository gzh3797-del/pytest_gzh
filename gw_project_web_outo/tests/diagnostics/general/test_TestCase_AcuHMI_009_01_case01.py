import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_diagnostics(page, submenu: str):
    if "/diagnostics" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/protocols",
            "/maintenance", "/templates", "/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Diagnostics").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_01_case01
# 用例标题：查看网络状态，点击Refresh，网络接口信息正确显示
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Diagnostics->Network Status
#   2. 点击Refresh按钮
# 预期结果：
#   2. 网络接口状态信息正确显示（接口名称、IP地址、状态等）
def test_TestCase_AcuHMI_009_01_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_diagnostics(page, "Network Status")

    page.get_by_role("button", name="Refresh").click()
    page.wait_for_timeout(2000)

    # Verify network info is displayed (plain-text pre blocks, not an HTML table)
    pre_blocks = page.locator("pre")
    assert pre_blocks.count() > 0 and pre_blocks.first.inner_text().strip() != "",         "刷新后网络接口信息应显示，但内容为空"
