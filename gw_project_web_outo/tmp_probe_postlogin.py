"""Probe post-login page structure to understand navigation."""
import sys
sys.path.insert(0, r"C:\AI工具\autotest\gw_project_web_outo")

from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, slow_mo=100)
    ctx = browser.new_context(ignore_https_errors=True, base_url=BASE_URL)
    page = ctx.new_page()

    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
    page.get_by_role("textbox", name="Enter User Name").press("Tab")
    page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    try:
        page.get_by_role("button", name="Cancel").click(timeout=3000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    print("Post-login URL:", page.url)
    page.screenshot(path="screenshots/tmp_postlogin.png")

    # Check header spans
    header_spans = page.locator("header span")
    print(f"\nHeader spans ({header_spans.count()}):")
    for i in range(min(20, header_spans.count())):
        txt = header_spans.nth(i).inner_text().strip()
        if txt:
            print(f"  {i}: {txt!r}")

    # Check left nav items
    nav_items = page.locator(".left-nav-item")
    print(f"\nLeft nav items ({nav_items.count()}):")
    for i in range(nav_items.count()):
        txt = nav_items.nth(i).inner_text().strip()[:60]
        print(f"  {i}: {txt!r}")

    # Check all top-level navigation
    print("\nAll roles 'navigation':")
    navs = page.get_by_role("navigation")
    print(f"  count: {navs.count()}")

    # Check menuitem
    menu_items = page.get_by_role("menuitem")
    print(f"\nMenu items ({menu_items.count()}):")
    for i in range(min(10, menu_items.count())):
        txt = menu_items.nth(i).inner_text().strip()
        print(f"  {i}: {txt!r}")

    # Look for the side nav
    sidebar = page.locator(".el-menu, .sidebar, .side-nav, nav")
    print(f"\nSidebar elements (.el-menu etc): {sidebar.count()}")

    # Check for any link in header
    header = page.locator("header")
    print(f"\nHeader HTML (first 500 chars):")
    try:
        html = header.first.inner_html()[:500]
        print(html)
    except Exception as e:
        print(f"Error: {e}")

    ctx.close()
    browser.close()
