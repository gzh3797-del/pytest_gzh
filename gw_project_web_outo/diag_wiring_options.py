"""Diagnostic: list all Wiring Configuration options."""
from pages.login_page import LoginPage


def _nav_to_templates(page):
    if "/#/templates" not in page.url:
        if not any(s in page.url for s in [
            "/#/systemSettings", "/#/templates", "/#/protocols",
            "/#/maintenance", "/#/diagnostics", "/#/userManagement", "/#/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_first_visible(page):
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    return False


def test_diag_wiring_options(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Select Typical Model first (required before Wiring Config)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    print("\n=== Typical Model options ===")
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible():
                print(f"  '{item.inner_text().strip()}'")
        except Exception:
            pass
    _click_first_visible(page)

    # Wiring Configuration options
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    print("\n=== Wiring Configuration options ===")
    wiring_options = []
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible():
                txt = item.inner_text().strip()
                wiring_options.append(txt)
                print(f"  '{txt}'")
        except Exception:
            pass

    page.keyboard.press("Escape")
    print(f"\nTotal wiring options: {len(wiring_options)}")
    assert True
