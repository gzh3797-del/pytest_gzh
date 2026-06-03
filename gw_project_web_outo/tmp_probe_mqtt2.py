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

    print("URL:", page.url)

    # Check current MQTT Enable state
    enable_radio = page.locator(".el-form-item").filter(has_text="MQTT Enable")
    print("MQTT Enable section text:", repr(enable_radio.inner_text()))

    # Check if Enable radio selected
    enable_btn = page.locator(".el-form-item").filter(has_text="MQTT Enable").locator(".el-radio").filter(has_text="Enable")
    disable_btn = page.locator(".el-form-item").filter(has_text="MQTT Enable").locator(".el-radio").filter(has_text="Disable")

    is_enabled = 'is-checked' in (enable_btn.get_attribute("class") or "")
    print("Currently enabled:", is_enabled)

    if not is_enabled:
        print("Clicking Enable...")
        enable_btn.locator(".el-radio__inner").click()
        page.wait_for_timeout(1000)

    print("\n--- After enabling MQTT ---")
    print("--- el-form-item labels ---")
    items = page.locator(".el-form-item__label").all()
    for item in items:
        print("  LABEL:", repr(item.inner_text()))

    print("\n--- input fields ---")
    inputs = page.locator("input").all()
    for inp in inputs:
        placeholder = inp.get_attribute("placeholder") or ""
        name = inp.get_attribute("name") or ""
        type_ = inp.get_attribute("type") or "text"
        value = inp.input_value() if type_ not in ('radio', 'checkbox') else ""
        print(f"  type={type_} placeholder={placeholder!r} name={name!r} value={value!r}")

    print("\n--- buttons ---")
    btns = page.get_by_role("button").all()
    for btn in btns:
        print("  BTN:", repr(btn.inner_text()))

    browser.close()
