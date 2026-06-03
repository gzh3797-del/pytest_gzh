"""Probe Post Channel 1 page structure."""
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

    # Navigate to Post Channel 1
    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Post Channels").click()
    page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Post Channel 1").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    print("URL:", page.url)

    page.screenshot(path="screenshots/tmp_postchannel1_page.png")

    # Tabs
    tabs = page.locator("[role='tab'], .el-tabs__item")
    print(f"\nTabs ({tabs.count()}):")
    for i in range(tabs.count()):
        print(f"  {i}: {tabs.nth(i).inner_text().strip()!r}")

    # Buttons
    buttons = page.get_by_role("button")
    print(f"\nButtons ({buttons.count()}):")
    for i in range(min(20, buttons.count())):
        txt = buttons.nth(i).inner_text().strip()[:50]
        if txt:
            print(f"  {i}: {txt!r}")

    # Form items
    form_items = page.locator(".el-form-item")
    print(f"\nForm items ({form_items.count()}):")
    for i in range(min(20, form_items.count())):
        txt = form_items.nth(i).inner_text().strip()[:100]
        if txt:
            print(f"  {i}: {txt!r}")

    # Radio elements
    radios = page.locator(".el-radio, .el-radio-button")
    print(f"\nRadio elements ({radios.count()}):")
    for i in range(min(10, radios.count())):
        print(f"  {i}: {radios.nth(i).inner_text().strip()!r}")

    ctx.close()
    browser.close()
