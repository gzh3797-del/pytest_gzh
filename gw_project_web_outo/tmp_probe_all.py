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


def probe_page(page, protocol, sub=None):
    _nav_protocol(page, protocol, sub)
    page.wait_for_timeout(800)
    print(f"\n=== {protocol} {sub or ''} ===")
    print("URL:", page.url)
    print("Form labels:")
    for item in page.locator(".el-form-item__label").all():
        print(f"  {repr(item.inner_text())}")
    print("Input fields:")
    for inp in page.locator("input[type='text'], input[type='number']").all():
        ph = inp.get_attribute("placeholder") or ""
        val = ""
        try:
            val = inp.input_value()
        except:
            pass
        print(f"  placeholder={ph!r} value={val!r}")
    print("Buttons:")
    for btn in page.get_by_role("button").all():
        txt = btn.inner_text().strip()
        if txt:
            print(f"  BTN: {repr(txt)}")
    # Check enable state
    enable_items = page.locator(".el-form-item").filter(has_text="Enable")
    print(f"Enable items: {enable_items.count()}")
    for item in enable_items.all():
        label = item.locator(".el-form-item__label")
        if label.count() > 0:
            radios = item.locator(".el-radio").all()
            for r in radios:
                cls = r.get_attribute("class") or ""
                print(f"  [{label.inner_text()}] radio={repr(r.inner_text())} checked={'is-checked' in cls}")


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

    probe_page(page, "BACnet/IP")
    probe_page(page, "AWS IoT")
    probe_page(page, "Azure IoT")

    browser.close()
