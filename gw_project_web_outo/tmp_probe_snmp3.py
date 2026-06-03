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

    # Click the SNMP Version dropdown
    ver_item = page.locator(".el-form-item").filter(has_text="SNMP Version")
    ver_select = ver_item.locator(".el-select")
    print("el-select HTML snippet:", ver_select.get_attribute("class"))
    print("el-select inner:", repr(ver_select.inner_text()))

    # Click to open
    ver_select.click()
    page.wait_for_timeout(500)

    # Find options
    options = page.locator(".el-select-dropdown__item").all()
    print(f"\nDropdown options ({len(options)}):")
    for opt in options:
        print(f"  {repr(opt.inner_text())} class={repr(opt.get_attribute('class'))}")

    # Select v2c
    v2c_opt = None
    for opt in options:
        txt = opt.inner_text()
        if 'v2' in txt.lower() or 'v2c' in txt.lower() or '2c' in txt.lower():
            v2c_opt = opt
            break

    if v2c_opt:
        print(f"\nClicking option: {repr(v2c_opt.inner_text())}")
        v2c_opt.click()
        page.wait_for_timeout(500)

        print("\nAfter selecting v2c - form labels:")
        for item in page.locator(".el-form-item__label").all():
            print(f"  LABEL: {repr(item.inner_text())}")
        print("\nInput fields:")
        for inp in page.locator("input[type='text'], input[type='number']").all():
            print(f"  placeholder={inp.get_attribute('placeholder')!r} value={inp.input_value()!r}")

    browser.close()
