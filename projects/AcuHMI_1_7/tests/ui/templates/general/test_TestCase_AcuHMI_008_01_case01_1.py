import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_templates(page, submenu="Template List"):
    if "/templates" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/protocols",
            "/maintenance", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_008_01_case01_1
# 用例标题：查看Official是否只有PXM350等产品的模板（用户确认不实现自动化）
@pytest.mark.skip(reason="用户确认不实现自动化：Official模板列表仅查看，不适合自动化")
def test_TestCase_AcuHMI_008_01_case01_1(login_page: LoginPage):
    pass
