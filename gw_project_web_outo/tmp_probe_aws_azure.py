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

    # --- AWS IoT after enabling ---
    _nav_protocol(page, "AWS IoT")
    page.wait_for_timeout(800)
    print("=== AWS IoT ===")

    # Enable AWS IoT
    aws_item = page.locator(".el-form-item").filter(has_text="AWS IoT Enable")
    en_radio = aws_item.locator(".el-radio").filter(has_text="Enable")
    if 'is-checked' not in (en_radio.get_attribute("class") or ""):
        en_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)
        print("AWS IoT enabled. Fields:")
        for item in page.locator(".el-form-item__label").all():
            print(f"  LABEL: {repr(item.inner_text())}")
        for inp in page.locator("input[type='text'], input[type='number']").all():
            print(f"  ph={inp.get_attribute('placeholder')!r} val={inp.input_value()!r}")
        print("Buttons:")
        for btn in page.get_by_role("button").all():
            txt = btn.inner_text().strip()
            if txt:
                print(f"  BTN: {repr(txt)}")

    # --- Azure IoT Interval field ---
    _nav_protocol(page, "Azure IoT")
    page.wait_for_timeout(800)
    print("\n=== Azure IoT ===")
    for item in page.locator(".el-form-item__label").all():
        print(f"  LABEL: {repr(item.inner_text())}")
    for inp in page.locator("input").all():
        typ = inp.get_attribute("type") or ""
        ph = inp.get_attribute("placeholder") or ""
        val = ""
        try:
            val = inp.input_value()
        except:
            pass
        print(f"  type={typ!r} ph={ph!r} val={val!r}")
    print("Buttons:")
    for btn in page.get_by_role("button").all():
        txt = btn.inner_text().strip()
        if txt:
            print(f"  BTN: {repr(txt)}")

    browser.close()
