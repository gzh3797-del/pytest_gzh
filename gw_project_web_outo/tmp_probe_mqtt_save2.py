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

    # Enable MQTT
    enable_item = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
    if 'is-checked' not in (enable_radio.get_attribute("class") or ""):
        enable_radio.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    # Fill ALL fields
    broker_field = page.get_by_placeholder("Enter Broker Address")
    port_field = page.get_by_placeholder("Enter Broker Port")
    clientid_field = page.get_by_placeholder("Enter Client ID")
    keepalive_field = page.get_by_placeholder("Enter Keep Alive")
    timeout_field = page.get_by_placeholder("Enter Timeout")

    broker_field.fill("test.broker.com")
    port_field.fill("1883")
    clientid_field.fill("test-client-001")
    keepalive_field.fill("60")
    timeout_field.fill("30")

    page.get_by_role("button", name="Save").click()
    print("Saved with all fields filled (valid domain):")
    try:
        msg = page.locator(".el-message").first
        msg.wait_for(state="visible", timeout=5000)
        print("  Message:", repr(msg.inner_text()), "class:", msg.get_attribute("class"))
    except Exception as e:
        print("  No toast message within 5s")
        errors = page.locator(".el-form-item__error").all()
        for err in errors:
            print("  Error:", repr(err.inner_text()))

    page.wait_for_timeout(2000)

    # Now test with invalid broker
    broker_field.fill("!@#$invalid")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    print("\nSaved with invalid broker:")
    errors = page.locator(".el-form-item__error").all()
    for err in errors:
        print("  Error:", repr(err.inner_text()))

    # Test with IP address
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    print("\n\nAfter reload - current field values:")
    fields = page.locator("input[type='text']").all()
    for f in fields:
        print(f"  placeholder={f.get_attribute('placeholder')!r} value={f.input_value()!r}")

    browser.close()
