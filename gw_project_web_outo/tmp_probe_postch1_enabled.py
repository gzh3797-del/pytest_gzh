"""Probe Post Channel 1 page when enabled with FTP."""
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
    page.wait_for_timeout(300)
    page.get_by_role("menuitem", name="Post Channel 1").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Enable Post Channel 1
    page.locator(".el-radio").filter(has_text="Enable").click()
    page.wait_for_timeout(500)
    print("After enabling URL:", page.url)

    page.screenshot(path="screenshots/tmp_postch1_enabled.png")

    # Form items after enable
    form_items = page.locator(".el-form-item")
    print(f"\nForm items after enable ({form_items.count()}):")
    for i in range(form_items.count()):
        txt = form_items.nth(i).inner_text().strip()[:120]
        if txt:
            print(f"  {i}: {txt!r}")

    # Buttons
    buttons = page.get_by_role("button")
    print(f"\nButtons ({buttons.count()}):")
    for i in range(min(20, buttons.count())):
        txt = buttons.nth(i).inner_text().strip()[:50]
        if txt:
            print(f"  {i}: {txt!r}")

    # Select FTP and see what appears
    print("\n--- Selecting FTP ---")
    try:
        page.locator(".el-form-item").filter(has_text="Post Method").locator(".el-select").click()
        page.wait_for_timeout(300)
        options = page.get_by_role("option")
        print(f"Options ({options.count()}):")
        for i in range(options.count()):
            print(f"  {options.nth(i).inner_text()!r}")
        page.get_by_role("option").first.click()
        page.wait_for_timeout(500)
    except Exception as e:
        print(f"Error selecting method: {e}")

    # Form items with FTP selected
    form_items2 = page.locator(".el-form-item")
    print(f"\nForm items with method selected ({form_items2.count()}):")
    for i in range(form_items2.count()):
        txt = form_items2.nth(i).inner_text().strip()[:120]
        if txt:
            print(f"  {i}: {txt!r}")

    # After Enable, check for Clear button
    buttons2 = page.get_by_role("button")
    print(f"\nButtons with FTP ({buttons2.count()}):")
    for i in range(min(20, buttons2.count())):
        txt = buttons2.nth(i).inner_text().strip()[:50]
        if txt:
            print(f"  {i}: {txt!r}")

    ctx.close()
    browser.close()
