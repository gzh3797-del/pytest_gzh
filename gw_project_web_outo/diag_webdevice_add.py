"""Diagnostic pytest test: inspect Web Device add dialog state."""
import pytest
from pages.login_page import LoginPage


def test_diag_webdevice(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Navigate to Web Devices
    if "/#/webDevice" not in page.url:
        try:
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        except Exception as e:
            print(f"Nav to Devices header: {e}")
        page.locator(".left-nav-item").filter(has_text="Web Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    print(f"\nCurrent URL: {page.url}")

    # Click Add Device
    page.get_by_role("button", name="Add Device").click()
    page.wait_for_timeout(1000)

    dialog = page.locator(".el-dialog")
    print(f"\nDialog visible: {dialog.is_visible()}")

    if dialog.count() > 0:
        # Print all form items
        form_items = dialog.locator(".el-form-item")
        print(f"\nForm items count: {form_items.count()}")
        for j in range(form_items.count()):
            item = form_items.nth(j)
            label_els = item.locator(".el-form-item__label")
            label = label_els.inner_text() if label_els.count() > 0 else "(no label)"
            inputs = item.locator("input")
            inp_count = inputs.count()
            print(f"  Item {j}: label='{label}', inputs={inp_count}")
            for k in range(inp_count):
                inp = inputs.nth(k)
                print(f"    input[{k}]: type={inp.get_attribute('type')}, "
                      f"placeholder='{inp.get_attribute('placeholder')}', "
                      f"value='{inp.input_value()}'")

        # Fill fields
        def fill_by_label(text, value):
            inp = dialog.locator(".el-form-item").filter(has_text=text).locator("input").first
            inp.fill(value)
            print(f"  Filled '{text}' = '{value}'")

        fill_by_label("Device Name", "TestDev001")
        fill_by_label("Serial Number", "SN000001")
        fill_by_label("Model", "M001")

        # Try URL field
        url_input = dialog.locator("input[placeholder='---Enter URL---']")
        print(f"\nURL input count: {url_input.count()}")
        if url_input.count() > 0:
            url_input.fill("http://10.0.1.1")
            print("Filled URL: http://10.0.1.1")
        page.wait_for_timeout(500)

        # Print current values before confirm
        print("\nValues before Confirm:")
        all_inputs = dialog.locator("input")
        for k in range(all_inputs.count()):
            inp = all_inputs.nth(k)
            try:
                val = inp.input_value()
                ph = inp.get_attribute("placeholder") or ""
                print(f"  input[{k}]: value='{val}', placeholder='{ph}'")
            except Exception:
                pass

        # Click Confirm
        dialog.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(2500)

        # Check state after
        print(f"\nDialog still visible: {dialog.is_visible()}")
        print(f"Dialog count: {dialog.count()}")

        # Check for form errors
        errors = dialog.locator(".el-form-item__error")
        print(f"Form errors: {errors.count()}")
        for k in range(errors.count()):
            print(f"  Error [{k}]: '{errors.nth(k).inner_text()}'")

        # Check toasts
        toasts = page.locator(".el-message")
        print(f"\nToasts: {toasts.count()}")
        for k in range(toasts.count()):
            try:
                print(f"  Toast [{k}]: '{toasts.nth(k).inner_text()}'")
            except Exception:
                pass

        # Print dialog inner text
        try:
            print(f"\nDialog inner text (first 1000 chars):\n{dialog.inner_text()[:1000]}")
        except Exception as e:
            print(f"Dialog closed: {e}")

    # Confirm we see it as pass/fail
    assert True, "Diagnostic complete, check printed output"
