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

    # Navigate to MQTT section
    _nav_protocol(page, "MQTT", "General")
    page.wait_for_timeout(500)

    # Get all MQTT sub-menu items
    print("=== MQTT Sub-menu Items ===")
    mqtt_menu = page.locator(".el-submenu").filter(has_text="MQTT").last
    if mqtt_menu.count() == 0:
        mqtt_menu = page.locator(".el-menu-item-group").filter(has_text="MQTT")

    # Just get all visible menu items
    all_menuitems = page.get_by_role("menuitem").all()
    for mi in all_menuitems:
        txt = mi.inner_text().strip()
        if txt:
            print(f"  menuitem: {repr(txt)}")

    print("\nURL:", page.url)

    # Try to navigate to SSL/TLS
    print("\nTrying 'SSL/TLS':")
    try:
        page.get_by_role("menuitem", name="SSL/TLS").click(timeout=3000)
        page.wait_for_timeout(300)
        print("  URL after SSL/TLS:", page.url)
    except Exception as e:
        print("  FAILED:", e)

    print("\nTrying 'SSL':")
    _nav_protocol(page, "MQTT", "General")
    try:
        page.get_by_role("menuitem", name="SSL").click(timeout=3000)
        page.wait_for_timeout(300)
        print("  URL after SSL:", page.url)
    except Exception as e:
        print("  FAILED:", e)

    browser.close()
