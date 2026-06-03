"""Diagnostic: find the Add/Create entry for custom templates."""
import pytest
from pages.login_page import LoginPage


def _nav_to_template_list(page):
    if "/#/templates" not in page.url:
        if not any(s in page.url for s in [
            "/#/systemSettings", "/#/templates", "/#/protocols",
            "/#/maintenance", "/#/diagnostics", "/#/userManagement",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    item = page.locator(".el-menu-item").filter(has_text="Template List")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_templates4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_template_list(page)
    page.wait_for_timeout(500)

    # Scroll to top
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(300)

    # Print full page text (first 3000 chars)
    body_txt = page.locator("body").inner_text()
    print(f"\n=== Page text (first 2000 chars) ===\n{body_txt[:2000]}")

    # Find all elements containing "Custom" text
    print("\n=== Elements containing 'Custom' ===")
    custom_els = page.get_by_text("Custom", exact=False).all()
    for el in custom_els:
        try:
            html = el.evaluate("el => el.outerHTML")[:300]
            print(f"  text='{el.inner_text().strip()[:50]}' html={html}")
        except Exception as e:
            print(f"  ERROR: {e}")

    # Check all buttons including outside viewport — scroll through page
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(300)
    print("\n=== All buttons after scroll to bottom ===")
    all_btns = page.locator("button").all()
    print(f"  Total button count: {len(all_btns)}")
    for i, btn in enumerate(all_btns):
        try:
            txt = btn.inner_text().strip()
            cls = btn.get_attribute("class") or ""
            aria = btn.get_attribute("aria-label") or ""
            visible = btn.is_visible()
            if "el-button" in cls:
                print(f"  [{i}] text='{txt}' aria='{aria}' class={cls[:80]} visible={visible}")
        except Exception:
            pass

    # Scroll back to top, check section headers
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(300)

    # Check card/section structure
    print("\n=== el-card sections ===")
    cards = page.locator(".el-card, .section-header, [class*='card'], [class*='section']").all()
    for c in cards[:5]:
        try:
            if c.is_visible():
                txt = c.inner_text().strip()[:100]
                print(f"  '{txt}'")
        except Exception:
            pass

    # Check if there is an "Add" button with icon near "Custom"
    print("\n=== Headings/titles on page ===")
    headings = page.locator("h1, h2, h3, h4, .title, .el-card__header").all()
    for h in headings:
        try:
            if h.is_visible():
                print(f"  '{h.inner_text().strip()[:80]}'")
        except Exception:
            pass

    # Try "New Typical Energy Meter Template" menu item
    print("\n=== Clicking 'New Typical Energy Meter Template' ===")
    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    if ntet.count() > 0 and ntet.first.is_visible():
        ntet.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        print(f"  URL: {page.url}")
        # Print page buttons
        btns2 = page.locator("button").all()
        for btn in btns2:
            try:
                if btn.is_visible():
                    txt = btn.inner_text().strip()
                    cls = btn.get_attribute("class") or ""
                    if txt or "el-button" in cls:
                        print(f"  button text='{txt}' class={cls[:60]}")
            except Exception:
                pass
        # Print form items
        fis = page.locator(".el-form-item").all()
        for fi in fis:
            try:
                lbl = fi.locator(".el-form-item__label").first.inner_text().strip()
                print(f"  form item: '{lbl}'")
            except Exception:
                pass

    assert True
