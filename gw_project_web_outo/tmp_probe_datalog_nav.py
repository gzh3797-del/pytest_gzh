"""Probe Data Log navigation structure."""
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

    # Navigate to Data Log
    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    print("Data Log submenus:")
    menu_items = page.get_by_role("menuitem")
    for i in range(menu_items.count()):
        txt = menu_items.nth(i).inner_text().strip()
        if txt:
            print(f"  {txt!r}")

    # Now let's click on each one to see what's there
    print("\n\nPost Historical Data page structure:")
    page.get_by_role("menuitem", name="Post Historical Data").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    page.screenshot(path="screenshots/tmp_postchannel_detailed.png")

    # Get page HTML structure
    content = page.content()
    # Extract relevant portions (tabs, form items, buttons)
    import re
    # Find tabs
    tab_matches = re.findall(r'role="tab"[^>]*>([^<]+)', content)
    print("Tabs found in HTML:", tab_matches[:10])

    # Find all .el-tabs__item elements
    tab_items = page.locator(".el-tabs__item")
    print(f"\nel-tabs__item ({tab_items.count()}):")
    for i in range(tab_items.count()):
        print(f"  {i}: {tab_items.nth(i).inner_text().strip()!r}")

    # Check for any nav items or sub-navigation
    nav_items = page.locator(".nav-item, .sub-nav, li[class*='nav'], li[class*='menu']")
    print(f"\nNav items ({nav_items.count()}):")

    ctx.close()
    browser.close()
