"""Diagnostic: dump all menu items visible after entering System Settings."""
from pages.login_page import LoginPage


def test_diag_system_settings_menu(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Navigate to System Settings via dateTime URL
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    print(f"\nURL: {page.url}")

    # Dump ALL el-menu-item elements (visible or not)
    print("\n=== All .el-menu-item elements ===")
    for i, item in enumerate(page.locator(".el-menu-item").all()):
        try:
            txt = item.inner_text().strip()
            visible = item.is_visible()
            cls = item.get_attribute("class") or ""
            print(f"  [{i}] visible={visible} text='{txt}' class='{cls[:60]}'")
        except Exception as e:
            print(f"  [{i}] error: {e}")

    # Dump ALL li[role=menuitem] elements
    print("\n=== All li[role=menuitem] elements ===")
    for i, item in enumerate(page.locator("li[role='menuitem']").all()):
        try:
            txt = item.inner_text().strip()
            visible = item.is_visible()
            print(f"  [{i}] visible={visible} text='{txt}'")
        except Exception as e:
            print(f"  [{i}] error: {e}")

    # Dump left nav items
    print("\n=== .left-nav-item elements ===")
    for item in page.locator(".left-nav-item").all():
        try:
            txt = item.inner_text().strip()
            visible = item.is_visible()
            print(f"  visible={visible} text='{txt}'")
        except Exception:
            pass

    assert True
