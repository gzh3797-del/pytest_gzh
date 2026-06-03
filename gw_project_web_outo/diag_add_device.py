"""Diagnostic: explore Add Device form UI structure."""
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


def test_diag_add_device(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_physical_devices(page)
    print(f"\nURL after nav: {page.url}")

    # Click Add Device button
    add_btn = page.get_by_role("button", name="Add Device")
    print(f"\nAdd Device button count: {add_btn.count()}")
    if add_btn.count() == 0:
        # Try alternate selectors
        for txt in ["Add", "New Device", "添加"]:
            cnt = page.get_by_role("button", name=txt).count()
            print(f"  button '{txt}': {cnt}")
        for cls in [".el-button--primary", ".add-device-btn"]:
            cnt = page.locator(cls).count()
            print(f"  {cls}: {cnt}")
    else:
        add_btn.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print(f"\nURL after click: {page.url}")

        # Print all form items
        print("\n=== Form items on Add Device page ===")
        for fi in page.locator(".el-form-item").all():
            try:
                label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
                has_input = fi.locator("input").count() > 0
                has_select = fi.locator(".el-select").count() > 0
                has_textarea = fi.locator("textarea").count() > 0
                print(f"  label='{label}' input={has_input} select={has_select} textarea={has_textarea}")
            except Exception:
                pass

        # Print all buttons
        print("\n=== Buttons on Add Device page ===")
        for btn in page.get_by_role("button").all():
            try:
                txt = btn.inner_text().strip()
                if txt:
                    print(f"  '{txt}'")
            except Exception:
                pass

        # Try Protocol dropdown
        print("\n=== Protocol dropdown options ===")
        proto_fi = page.locator(".el-form-item").filter(has_text="Protocol")
        print(f"  Protocol form items: {proto_fi.count()}")
        if proto_fi.count() > 0:
            proto_fi.first.locator(".el-select").first.click()
            page.wait_for_timeout(300)
            for opt in page.locator(".el-select-dropdown__item").all():
                try:
                    if opt.is_visible():
                        print(f"  option: '{opt.inner_text().strip()}'")
                except Exception:
                    pass
            page.keyboard.press("Escape")
            page.wait_for_timeout(200)

    assert True
