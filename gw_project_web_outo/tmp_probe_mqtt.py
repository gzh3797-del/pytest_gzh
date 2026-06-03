import pytest
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL
import sys
sys.stdout.reconfigure(encoding='utf-8')


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

    # Login
    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    from pages.login_page import LoginPage
    lp = LoginPage(page)
    lp.login()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # Nav to MQTT General
    _nav_protocol(page, "MQTT", "General")
    page.wait_for_timeout(1000)

    print("URL:", page.url)
    print("\n--- el-form-item labels ---")
    items = page.locator(".el-form-item__label").all()
    for item in items:
        print("  LABEL:", repr(item.inner_text()))

    print("\n--- input/textarea fields ---")
    inputs = page.locator("input, textarea").all()
    for inp in inputs:
        placeholder = inp.get_attribute("placeholder") or ""
        name = inp.get_attribute("name") or ""
        id_ = inp.get_attribute("id") or ""
        type_ = inp.get_attribute("type") or "text"
        print(f"  type={type_} name={name!r} id={id_!r} placeholder={placeholder!r}")

    print("\n--- buttons ---")
    btns = page.get_by_role("button").all()
    for btn in btns:
        print("  BTN:", repr(btn.inner_text()))

    browser.close()
