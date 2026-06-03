"""Diagnostic: inspect Parameters dropdown options in Batch Update dialog."""
from pages.login_page import LoginPage


def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_params_dropdown(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "BACnet/IP")

    # Enable BACnet
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    if "is-checked" not in (bacnet_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Open Parameter Config for AcuRev4100
    device_table_fi = None
    for fi in page.locator(".el-form-item").all():
        if fi.locator(".el-switch, .el-checkbox").count() > 3:
            device_table_fi = fi
            break

    acurev_row = device_table_fi.locator("tbody tr").filter(has_text="AcuRev4100").first
    acurev_row.locator(".el-button--primary").first.click()
    page.wait_for_timeout(1000)

    param_dialog = page.locator(".el-dialog").filter(has_text="Parameter Config").first

    # Click COV Batch Update button
    param_dialog.get_by_role("button").filter(has_text="COV Batch Update").first.click()
    page.wait_for_timeout(1000)

    batch_dialog = page.locator(".el-dialog").filter(has_text="Batch Update").last

    # Click Parameters dropdown and inspect
    params_select = batch_dialog.locator(".el-form-item").filter(has_text="Parameters").first.locator(".el-select").first
    params_select.click()
    page.wait_for_timeout(800)

    print("\n=== After clicking Parameters dropdown ===")
    print(f"  el-select-dropdown count: {page.locator('.el-select-dropdown').count()}")
    print(f"  el-select-dropdown visible: {page.locator('.el-select-dropdown:visible').count()}")

    # All select dropdowns
    for i, dd in enumerate(page.locator(".el-select-dropdown").all()):
        try:
            visible = dd.is_visible()
            txt = dd.inner_text()[:200]
            print(f"  dropdown[{i}] visible={visible}: '{txt}'")
        except Exception:
            pass

    # Check for virtual list items (el-select uses virtual scrolling sometimes)
    print(f"\n  .el-select-dropdown__item count: {page.locator('.el-select-dropdown__item').count()}")
    print(f"  .el-virtual-list__item count: {page.locator('.el-virtual-list__item').count()}")
    print(f"  .el-scrollbar__view li count: {page.locator('.el-scrollbar__view li').count()}")

    # Check the wrapper HTML to see aria-expanded
    wrapper_html = params_select.locator(".el-select__wrapper").evaluate("el => el.outerHTML")
    print(f"\n  Wrapper aria-expanded: {'aria-expanded=\"true\"' in wrapper_html}")

    # Try typing to filter
    print("\n=== Typing to search ===")
    search_input = params_select.locator("input.el-select__input")
    if search_input.count() > 0:
        search_input.fill("Phase")
        page.wait_for_timeout(500)
        print(f"  After typing 'Phase':")
        print(f"  el-select-dropdown__item count: {page.locator('.el-select-dropdown__item').count()}")
        for item in page.locator(".el-select-dropdown__item").all():
            try:
                if item.is_visible():
                    print(f"    '{item.inner_text().strip()}'")
            except Exception:
                pass

        # Clear search
        search_input.fill("")
        page.wait_for_timeout(300)
        print(f"  After clearing:")
        print(f"  el-select-dropdown__item count: {page.locator('.el-select-dropdown__item').count()}")
        for item in page.locator(".el-select-dropdown__item").all():
            try:
                if item.is_visible():
                    print(f"    '{item.inner_text().strip()}'")
            except Exception:
                pass

    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    assert True
