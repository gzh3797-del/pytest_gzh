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

    _nav_protocol(page, "MQTT", "General")
    page.wait_for_timeout(1000)

    print("Current state:")
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    print("  MQTT Enable section:", repr(enable_item.inner_text()))

    enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
    is_enabled = 'is-checked' in (enable_radio.get_attribute("class") or "")
    print("  Is enabled:", is_enabled)

    if is_enabled:
        print("  Current field values:")
        for inp in page.locator("input[type='text']").all():
            print(f"    {inp.get_attribute('placeholder')!r} = {inp.input_value()!r}")

        # Try modifying broker and saving
        broker = page.get_by_placeholder("Enter Broker Address")
        current = broker.input_value()
        print(f"\n  Current broker: {current!r}")

        broker.fill("test.broker.com")
        print("  After filling 'test.broker.com', clicking Save...")
        page.get_by_role("button", name="Save").click()

        try:
            msg = page.locator(".el-message").first
            msg.wait_for(state="visible", timeout=5000)
            print("  SUCCESS message:", repr(msg.inner_text()))
        except Exception as e:
            print("  No message:", e)
            errs = page.locator(".el-form-item__error").all()
            for err in errs:
                print("  Form error:", repr(err.inner_text()))

    else:
        print("  MQTT is disabled - need to enable first")
        # Enable it
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(300)

        print("  After enabling, field values:")
        for inp in page.locator("input[type='text']").all():
            print(f"    {inp.get_attribute('placeholder')!r} = {inp.input_value()!r}")

        # Generate Client ID
        gen_btn = page.get_by_role("button", name="Generate Client ID")
        if gen_btn.count() > 0:
            gen_btn.click()
            page.wait_for_timeout(300)
            print("  Client ID after generate:", page.get_by_placeholder("Enter Client ID").input_value())

        # Fill broker
        broker = page.get_by_placeholder("Enter Broker Address")
        broker.fill("test.broker.com")
        print("  Clicking Save...")
        page.get_by_role("button", name="Save").click()

        try:
            msg = page.locator(".el-message").first
            msg.wait_for(state="visible", timeout=5000)
            print("  SUCCESS message:", repr(msg.inner_text()))
        except Exception as e:
            print("  No message:", e)
            errs = page.locator(".el-form-item__error").all()
            for err in errs:
                print("  Form error:", repr(err.inner_text()))

    browser.close()
