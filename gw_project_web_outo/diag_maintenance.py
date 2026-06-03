"""Diagnostic: find Maintenance/Reboot entry from System Settings area."""
from pages.login_page import LoginPage


def test_diag_maintenance(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Must enter System Settings area first so left nav shows Maintenance
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    print(f"\nURL: {page.url}")

    print("\n=== .left-nav-item elements ===")
    for item in page.locator(".left-nav-item").all():
        try:
            txt = item.inner_text().strip()
            visible = item.is_visible()
            print(f"  visible={visible} '{txt}'")
        except Exception:
            pass

    # Click Maintenance
    maint = page.locator(".left-nav-item").filter(has_text="Maintenance").first
    assert maint.count() > 0, "未找到 Maintenance left-nav-item"
    maint.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    print(f"\nAfter Maintenance click URL: {page.url}")

    print("\n=== Sub menu items under Maintenance ===")
    for item in page.locator(".el-menu-item").all():
        try:
            txt = item.inner_text().strip()
            visible = item.is_visible()
            if txt:
                print(f"  visible={visible} '{txt}'")
        except Exception:
            pass

    print("\n=== All buttons ===")
    for btn in page.get_by_role("button").all():
        try:
            txt = btn.inner_text().strip()
            if txt:
                print(f"  '{txt}'")
        except Exception:
            pass

    print("\n=== Page text ===")
    try:
        print(page.locator("main, .el-main, .app-main").first.inner_text()[:600])
    except Exception:
        print(page.locator("body").inner_text()[:600])

    assert True
