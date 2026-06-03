"""Probe Post Channels page - full UI discovery."""
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

    # Navigate to Data Log first
    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Get all menuitem texts first
    menu_items = page.get_by_role("menuitem")
    print("All menu items:")
    for i in range(menu_items.count()):
        txt = menu_items.nth(i).inner_text().strip()
        print(f"  {i}: {txt!r}")

    print("\n--- Clicking 'Post Channels' ---")
    try:
        page.get_by_role("menuitem", name="Post Channels").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        print("URL after Post Channels click:", page.url)
    except Exception as e:
        print(f"Error clicking Post Channels: {e}")

    # Get all menu items again (might have expanded)
    menu_items2 = page.get_by_role("menuitem")
    print(f"\nMenu items after click ({menu_items2.count()}):")
    for i in range(menu_items2.count()):
        txt = menu_items2.nth(i).inner_text().strip()
        print(f"  {i}: {txt!r}")

    # Check full form structure
    form_items = page.locator(".el-form-item")
    print(f"\nAll form items ({form_items.count()}):")
    for i in range(form_items.count()):
        txt = form_items.nth(i).inner_text().strip()[:100]
        if txt:
            print(f"  {i}: {txt!r}")

    # Check if there are sections for Post Channel
    print("\nSearching for 'Post Channel' text in page:")
    post_ch_count = page.get_by_text("Post Channel", exact=False).count()
    print(f"  Found: {post_ch_count} elements")

    # Get all text on page
    page.screenshot(path="screenshots/tmp_postchannel_discovery.png")

    ctx.close()
    browser.close()
