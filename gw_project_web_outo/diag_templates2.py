"""Diagnostic: inspect Template List page buttons and form structure."""
import pytest
from pages.login_page import LoginPage


def _nav_to_template_list(page):
    if "/#/templates" not in page.url:
        if not any(s in page.url for s in [
            "/#/systemSettings", "/#/templates", "/#/protocols",
            "/#/maintenance", "/#/diagnostics", "/#/userManagement", "/#/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_templates2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_template_list(page)
    page.wait_for_timeout(800)
    print(f"\nURL: {page.url}")

    # Print all buttons with their outerHTML
    print("\n=== All buttons (outerHTML) ===")
    btns = page.locator("button").all()
    for i, btn in enumerate(btns):
        try:
            if btn.is_visible():
                html = btn.evaluate("el => el.outerHTML")
                txt = btn.inner_text().strip()
                label = btn.get_attribute("aria-label") or ""
                print(f"  [{i}] text='{txt}' aria-label='{label}' html={html[:200]}")
        except Exception as e:
            print(f"  [{i}] ERROR: {e}")

    # Check el-button elements
    print("\n=== el-button elements ===")
    el_btns = page.locator(".el-button").all()
    for i, btn in enumerate(el_btns):
        try:
            if btn.is_visible():
                txt = btn.inner_text().strip()
                label = btn.get_attribute("aria-label") or ""
                html = btn.evaluate("el => el.outerHTML")[:300]
                print(f"  [{i}] text='{txt}' aria-label='{label}'")
                print(f"       html: {html}")
        except Exception as e:
            print(f"  [{i}] ERROR: {e}")

    # Check tabs
    print("\n=== Tab labels ===")
    tab_labels = page.locator(".el-tabs__item, [role='tab']").all()
    for tl in tab_labels:
        try:
            print(f"  tab: '{tl.inner_text().strip()}' visible={tl.is_visible()}")
        except Exception:
            pass

    # Check table rows
    print("\n=== Table rows ===")
    rows = page.locator("tbody tr").all()
    print(f"  row count: {len(rows)}")
    for i, row in enumerate(rows[:3]):
        try:
            print(f"  row[{i}]: '{row.inner_text().strip()[:100]}'")
        except Exception:
            pass

    # Check page text keywords
    for kw in ["Add Template", "Create", "New Template", "Custom", "Import", "Template"]:
        count = page.get_by_text(kw, exact=False).count()
        print(f"  text '{kw}': {count} matches")

    assert True
