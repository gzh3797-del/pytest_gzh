"""Diagnostic: switch to TCP and explore form fields."""
from pages.login_page import LoginPage


def _nav_to_physical_devices(page):
    if "/#/physicalDevice" not in page.url or "addDevice" in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_visible_option(page, option_text: str = ""):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def test_diag_add_device3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_physical_devices(page)
    page.get_by_role("button", name="Add Device").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Click TCP radio
    page.locator(".el-radio").filter(has_text="TCP").click()
    page.wait_for_timeout(500)

    print("\n=== Form items after switching to TCP ===")
    for fi in page.locator(".el-form-item").all():
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
            has_input = fi.locator("input:not([type='radio'])").count() > 0
            has_select = fi.locator(".el-select").count() > 0
            print(f"  label='{label}' input={has_input} select={has_select}")
        except Exception:
            pass

    # Check Template options
    print("\n=== Template dropdown options ===")
    tmpl_fi = page.locator(".el-form-item").filter(has_text="Template")
    if tmpl_fi.count() > 0:
        tmpl_fi.first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        opts = []
        for opt in page.locator(".el-select-dropdown__item").all():
            try:
                if opt.is_visible():
                    txt = opt.inner_text().strip()
                    opts.append(txt)
            except Exception:
                pass
        print(f"  Total options: {len(opts)}")
        for o in opts[:5]:
            print(f"  '{o}'")
        if len(opts) > 5:
            print(f"  ... ({len(opts)-5} more)")
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)

    # Try to fill form and save
    print("\n=== Attempting to fill and save TCP device ===")
    # Device Name
    name_inp = page.locator(".el-form-item").filter(has_text="Device Name").first.locator("input").first
    name_inp.fill("DiagTCP_test")
    page.wait_for_timeout(100)

    # Select Template (first available)
    tmpl_fi = page.locator(".el-form-item").filter(has_text="Template")
    if tmpl_fi.count() > 0:
        tmpl_fi.first.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        _click_visible_option(page, "")

    # IP Address
    ip_fi = page.locator(".el-form-item").filter(has_text="IP")
    print(f"  IP form items: {ip_fi.count()}")
    if ip_fi.count() > 0:
        ip_inp = ip_fi.first.locator("input").first
        ip_inp.fill("192.168.99.99")
        page.wait_for_timeout(100)

    # Port
    port_fi = page.locator(".el-form-item").filter(has_text="Port")
    print(f"  Port form items: {port_fi.count()}")

    # Try save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    print(f"\n  URL after Save: {page.url}")
    errors = page.locator(".el-message--error, .el-form-item__error").all()
    for e in errors:
        try:
            if e.is_visible():
                print(f"  Error: '{e.inner_text()}'")
        except Exception:
            pass
    success = page.locator(".el-message--success").count()
    print(f"  Success messages: {success}")

    # Check if device appears in list
    if "physicalDevices" in page.url and "addDevice" not in page.url:
        row = page.locator("tbody tr").filter(has_text="DiagTCP_test")
        print(f"  Device in list: {row.count() > 0}")

    # Cancel and go back
    cancel = page.get_by_role("button", name="Cancel")
    if cancel.count() > 0:
        cancel.first.click()
        page.wait_for_timeout(500)

    assert True
