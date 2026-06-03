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

    # Try saving with a valid broker address
    broker_field = page.get_by_placeholder("Enter Broker Address")
    broker_field.fill("192.168.1.100")

    print("BEFORE save - checking for messages/errors:")
    print("  .el-message count:", page.locator(".el-message").count())

    page.get_by_role("button", name="Save").click()
    print("IMMEDIATELY after save - checking for messages:")
    msg = page.locator(".el-message").first
    print("  .el-message count:", page.locator(".el-message").count())

    # Try to catch the message quickly
    try:
        msg.wait_for(state="visible", timeout=5000)
        print("  Message appeared:", repr(msg.inner_text()))
        print("  Message class:", msg.get_attribute("class"))
    except Exception as e:
        print("  No message appeared within 5s:", e)

    page.wait_for_timeout(2000)
    print("AFTER 2s - .el-message count:", page.locator(".el-message").count())

    # Now try invalid value
    broker_field.fill("!@#$%^invalid")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)
    print("\nAfter invalid save:")
    print("  .el-form-item__error count:", page.locator(".el-form-item__error").count())
    errors = page.locator(".el-form-item__error").all()
    for e in errors:
        print("  Error text:", repr(e.inner_text()))
    print("  .el-message--error count:", page.locator(".el-message--error").count())

    browser.close()
