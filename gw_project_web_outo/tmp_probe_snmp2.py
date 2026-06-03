import sys
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL


def _nav_protocol(page, protocol, sub=None):
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


with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    ctx = browser.new_context(ignore_https_errors=True)
    page = ctx.new_page()

    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    from pages.login_page import LoginPage
    lp = LoginPage(page)
    lp.login()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    _nav_protocol(page, "SNMP")
    page.wait_for_timeout(1000)

    print("=== SNMP Page ===")
    print("URL:", page.url)

    print("\n--- SNMP Version selector ---")
    ver_item = page.locator(".el-form-item").filter(has_text="SNMP Version")
    print("Text:", repr(ver_item.inner_text()))

    # Check el-select for version
    ver_select = ver_item.locator(".el-select")
    if ver_select.count() > 0:
        print("El-select found:", repr(ver_select.inner_text()))
    # Check radio buttons for version
    ver_radios = ver_item.locator(".el-radio, .el-radio-button")
    print("Radios count:", ver_radios.count())
    for r in ver_radios.all():
        cls = r.get_attribute("class") or ""
        print(f"  radio: text={repr(r.inner_text())} checked={'is-checked' in cls}")

    # Try to change version to v2c
    print("\nAttempting to change version to v2c...")
    v2c_radio = ver_item.locator(".el-radio").filter(has_text="v2c")
    if v2c_radio.count() > 0:
        print("  v2c radio found")
        v2c_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)
        print("  After clicking v2c:")
        for inp in page.locator("input[type='text'], input[type='number']").all():
            print(f"    placeholder={inp.get_attribute('placeholder')!r} value={inp.input_value()!r}")
        for item in page.locator(".el-form-item__label").all():
            print(f"  LABEL: {repr(item.inner_text())}")
    else:
        # Maybe it's an el-select dropdown
        ver_select_input = ver_item.locator(".el-input__inner")
        if ver_select_input.count() > 0:
            print("  El-select input found, current value:", repr(ver_select_input.input_value()))
            ver_select_input.click()
            page.wait_for_timeout(300)
            options = page.locator(".el-select-dropdown .el-select-dropdown__item")
            print("  Dropdown options:", [o.inner_text() for o in options.all()])
            # Click v2c option
            v2c_opt = page.get_by_role("option", name="v2c")
            if v2c_opt.count() == 0:
                v2c_opt = page.locator(".el-select-dropdown__item").filter(has_text="v2c")
            if v2c_opt.count() > 0:
                v2c_opt.click()
                page.wait_for_timeout(500)
                print("  After selecting v2c:")
                for inp in page.locator("input[type='text'], input[type='number']").all():
                    print(f"    placeholder={inp.get_attribute('placeholder')!r} value={inp.input_value()!r}")
                for item in page.locator(".el-form-item__label").all():
                    print(f"  LABEL: {repr(item.inner_text())}")

    browser.close()
