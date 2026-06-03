"""Probe Post Historical Data page to understand UI structure."""
import sys
sys.path.insert(0, r"C:\AI工具\autotest\gw_project_web_outo")

from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, slow_mo=100)
    ctx = browser.new_context(ignore_https_errors=True, base_url=BASE_URL)
    page = ctx.new_page()

    # Login
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
    page.get_by_role("textbox", name="Enter User Name").press("Tab")
    page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    try:
        page.get_by_role("button", name="Cancel").click(timeout=3000)
    except Exception:
        pass
    print("After login:", page.url)

    # Navigate to Post Historical Data
    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Post Historical Data").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    print("Page URL:", page.url)

    # Take screenshot
    page.screenshot(path="screenshots/tmp_postchannel_page.png")
    print("Screenshot saved")

    # Get all tabs
    tabs = page.locator("[role='tab']")
    print(f"\nTabs ({tabs.count()}):")
    for i in range(tabs.count()):
        print(f"  {i}: {tabs.nth(i).inner_text()[:60]!r}")

    # Get all buttons
    buttons = page.get_by_role("button")
    print(f"\nButtons ({buttons.count()}):")
    for i in range(min(20, buttons.count())):
        txt = buttons.nth(i).inner_text().strip()[:40]
        if txt:
            print(f"  {i}: {txt!r}")

    # Get form items with "Enable"
    print("\nForm items with 'Enable':")
    enable_items = page.locator(".el-form-item").filter(has_text="Enable")
    print(f"  Count: {enable_items.count()}")

    # Check all radio elements
    radios = page.locator(".el-radio")
    print(f"\nRadio elements ({radios.count()}):")
    for i in range(min(10, radios.count())):
        print(f"  {i}: {radios.nth(i).inner_text().strip()!r}")

    # Check all radio buttons
    radio_btns = page.locator(".el-radio-button")
    print(f"\nRadio-button elements ({radio_btns.count()}):")
    for i in range(min(10, radio_btns.count())):
        print(f"  {i}: {radio_btns.nth(i).inner_text().strip()!r}")

    ctx.close()
    browser.close()
