"""Diagnostic: inspect Function and Address Format dropdown options on create page."""
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


def test_diag_block_options(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_templates(page)
    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Fill Typical Model & Wiring first (required for Block section)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    items = page.locator(".el-select-dropdown__item").all()
    for item in items:
        try:
            if item.is_visible():
                item.click()
                break
        except Exception:
            pass
    page.wait_for_timeout(200)

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    items = page.locator(".el-select-dropdown__item").all()
    for item in items:
        try:
            if item.is_visible():
                item.click()
                break
        except Exception:
            pass
    page.wait_for_timeout(200)

    # ── Function dropdown options ──────────────────────────────────────────
    print("\n=== Function dropdown options ===")
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    items = page.locator(".el-select-dropdown__item").all()
    for item in items:
        try:
            if item.is_visible():
                print(f"  '{item.inner_text().strip()}'")
        except Exception:
            pass

    # ── Address Format dropdown options ────────────────────────────────────
    print("\n=== Address Format dropdown options ===")
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    addr_fi = page.locator(".el-form-item").filter(has_text="Address Format").first
    if addr_fi.count() > 0:
        addr_fi.locator(".el-select").click()
        page.wait_for_timeout(400)
        items = page.locator(".el-select-dropdown__item").all()
        for item in items:
            try:
                if item.is_visible():
                    print(f"  '{item.inner_text().strip()}'")
            except Exception:
                pass
    else:
        print("  Address Format field NOT found")

    page.keyboard.press("Escape")
    assert True
