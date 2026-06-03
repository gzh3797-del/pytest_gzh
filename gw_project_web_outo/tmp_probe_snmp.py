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

    # SNMP
    _nav_protocol(page, "SNMP")
    page.wait_for_timeout(1000)

    print("=== SNMP Page ===")
    print("URL:", page.url)
    print("\nForm items:")
    for item in page.locator(".el-form-item__label").all():
        print("  LABEL:", repr(item.inner_text()))

    print("\nInputs:")
    for inp in page.locator("input[type='text'], input[type='number']").all():
        print(f"  placeholder={inp.get_attribute('placeholder')!r} value={inp.input_value()!r}")

    print("\nRadios:")
    for ri in page.locator(".el-radio").all():
        cls = ri.get_attribute("class") or ""
        print(f"  text={repr(ri.inner_text())} checked={'is-checked' in cls}")

    # Check if SNMP has enable/disable
    enable_items = page.locator(".el-form-item").filter(has_text="SNMP Enable")
    print(f"\nSNMP Enable items: {enable_items.count()}")
    if enable_items.count() > 0:
        print("  Text:", repr(enable_items.first.inner_text()))

    # Try enabling SNMP
    snmp_en = page.locator(".el-form-item").filter(has_text="SNMP Enable")
    if snmp_en.count() > 0:
        en_radio = snmp_en.locator(".el-radio").filter(has_text="Enable")
        if 'is-checked' not in (en_radio.get_attribute("class") or ""):
            print("\nEnabling SNMP...")
            en_radio.locator(".el-radio__inner").click()
            page.wait_for_timeout(500)

            print("\nAfter enabling, form items:")
            for item in page.locator(".el-form-item__label").all():
                print("  LABEL:", repr(item.inner_text()))

            print("\nInputs after enable:")
            for inp in page.locator("input[type='text'], input[type='number']").all():
                print(f"  placeholder={inp.get_attribute('placeholder')!r} value={inp.input_value()!r}")

    browser.close()
