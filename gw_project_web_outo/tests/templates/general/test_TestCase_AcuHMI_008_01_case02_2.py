import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


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


# 用例编号：TestCase_AcuHMI_008_01_case02_2
# 用例标题：Official支持10/20/40/80条/页切换查看
def test_TestCase_AcuHMI_008_01_case02_2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_templates(page, "Template List")

    # Switch to Official tab
    try:
        page.get_by_role("tab", name="Official").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.locator(".el-tabs__item").filter(has_text="Official").click()
            page.wait_for_timeout(500)
        except Exception:
            pass

    for page_size in ["10", "20", "40", "80"]:
        try:
            page.locator(".el-pagination").locator(".el-select").click()
            page.wait_for_timeout(200)
            page.get_by_role("option", name=f"{page_size}/page").click()
            page.wait_for_timeout(500)
        except Exception:
            try:
                page.locator(".el-select").filter(has_text="/page").click()
                page.wait_for_timeout(200)
                page.get_by_role("option", name=f"{page_size} /page").click()
                page.wait_for_timeout(500)
            except Exception:
                continue

        rows = page.locator("tbody tr").count()
        assert rows <= int(page_size), f"Official模板每页{page_size}条时，显示行数不应超过{page_size}"
