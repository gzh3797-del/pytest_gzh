import sys
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL


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

    # Navigate to Protocols
    page.locator("header span").filter(has_text="AcuHMI").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.locator(".left-nav-item").filter(has_text="Protocols").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # Click MQTT to expand
    page.get_by_role("menuitem", name="MQTT").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    print("URL after MQTT click:", page.url)

    print("\nAll menu items after MQTT click:")
    for mi in page.get_by_role("menuitem").all():
        txt = mi.inner_text().strip()
        cls = mi.get_attribute("class") or ""
        if txt:
            print(f"  {repr(txt)} — class: {cls[:50]}")

    # Now click General
    page.get_by_role("menuitem", name="General").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    print("\nURL after General click:", page.url)

    print("\nAll menu items after General click:")
    for mi in page.get_by_role("menuitem").all():
        txt = mi.inner_text().strip()
        cls = mi.get_attribute("class") or ""
        if txt:
            print(f"  {repr(txt)} — class: {cls[:50]}")

    # Try to click User Credential
    print("\nAttempting 'User Credential' click...")
    uc = page.get_by_role("menuitem", name="User Credential")
    print(f"  Count: {uc.count()}")
    for m in uc.all():
        print(f"  Found: {repr(m.inner_text())} class: {m.get_attribute('class')[:50]}")

    browser.close()
