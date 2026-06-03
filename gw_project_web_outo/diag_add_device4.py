"""Diagnostic: check Add to Logger options and device list delete button."""
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
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    for item in page.locator(".el-select-dropdown__item").all():
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


def test_diag_add_device4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_physical_devices(page)
    page.get_by_role("button", name="Add Device").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Switch to TCP
    page.locator(".el-radio").filter(has_text="TCP").click()
    page.wait_for_timeout(300)

    # Add to Logger options
    print("\n=== Add to Logger options ===")
    logger_fi = page.locator(".el-form-item").filter(has_text="Add to Logger").first
    if logger_fi.count() > 0:
        logger_fi.locator(".el-select").first.click()
        page.wait_for_timeout(300)
        for opt in page.locator(".el-select-dropdown__item").all():
            try:
                if opt.is_visible():
                    print(f"  '{opt.inner_text().strip()}'")
            except Exception:
                pass
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)

    # Fill all required fields and save
    import time
    ts = str(int(time.time()))[-6:]
    device_name = f"DiagTCP_{ts}"

    page.locator(".el-form-item").filter(has_text="Device Name").first.locator("input").first.fill(device_name)
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="Serial Number").first.locator("input").first.fill(f"SN{ts}")
    page.wait_for_timeout(100)

    # Template
    page.locator(".el-form-item").filter(has_text="Template").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    # IP Address
    page.locator(".el-form-item").filter(has_text="IP Address").first.locator("input").first.fill("192.168.99.99")
    page.wait_for_timeout(100)

    # Modbus ID
    page.locator(".el-form-item").filter(has_text="Modbus ID").first.locator("input").first.fill("1")
    page.wait_for_timeout(100)

    # Add to Logger (first option)
    page.locator(".el-form-item").filter(has_text="Add to Logger").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)

    print(f"\n  URL after Save: {page.url}")
    for e in page.locator(".el-message--error, .el-form-item__error").all():
        try:
            if e.is_visible():
                print(f"  Error: '{e.inner_text()}'")
        except Exception:
            pass
    print(f"  Success: {page.locator('.el-message--success').count() > 0}")

    # If navigated to list, check device and inspect delete button
    if "addDevice" not in page.url:
        print(f"\n  Save succeeded! Device should be in list.")
        row = page.locator("tbody tr").filter(has_text=device_name)
        print(f"  Device row count: {row.count()}")
        if row.count() > 0:
            print("\n=== Delete button structure in device row ===")
            btns = row.first.locator("button, .el-button").all()
            for b in btns:
                try:
                    txt = b.inner_text().strip()
                    cls = b.get_attribute("class") or ""
                    print(f"  btn: text='{txt}' class='{cls[:60]}'")
                except Exception:
                    pass

            # Try delete
            print("\n=== Attempting delete ===")
            danger_btn = row.first.locator(".el-button--danger")
            if danger_btn.count() > 0:
                danger_btn.first.click()
            else:
                row.first.locator("button").last.click()
            page.wait_for_timeout(500)

            # Confirm dialog
            for name in ["Yes, continue", "Yes", "Confirm", "确认"]:
                btn = page.get_by_role("button", name=name)
                if btn.count() > 0 and btn.first.is_visible():
                    print(f"  Confirm button: '{name}'")
                    btn.first.click()
                    page.wait_for_timeout(800)
                    break

            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
            remaining = page.locator("tbody tr").filter(has_text=device_name)
            print(f"  After delete, device count: {remaining.count()}")

    assert True
