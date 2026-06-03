"""Diagnostic: explore Templates page navigation and structure."""
import pytest
from pages.login_page import LoginPage


def test_diag_templates(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    page.wait_for_timeout(500)

    # Step 1: Click AcuHMI in header
    print(f"\nCurrent URL: {page.url}")
    header_btns = page.locator("header").locator("button, span, a").all()
    print("\n=== Header clickable elements ===")
    for el in header_btns:
        try:
            if el.is_visible():
                print(f"  '{el.inner_text().strip()}' tag={el.evaluate('el => el.tagName')}")
        except Exception:
            pass

    # Try clicking AcuHMI header button
    try:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        print(f"\nAfter AcuHMI click URL: {page.url}")
    except Exception as e:
        print(f"\nERROR clicking AcuHMI: {e}")

    # Step 2: Check left nav items
    print("\n=== Left nav items ===")
    left_nav = page.locator(".left-nav-item").all()
    for item in left_nav:
        try:
            if item.is_visible():
                print(f"  '{item.inner_text().strip()}'")
        except Exception:
            pass

    # Click Templates
    try:
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        print(f"\nAfter Templates click URL: {page.url}")
    except Exception as e:
        print(f"\nERROR clicking Templates: {e}")

    # Step 3: Check sub-menu structure
    print("\n=== Sub-menu titles (el-sub-menu__title) ===")
    sub_titles = page.locator("div.el-sub-menu__title").all()
    for st in sub_titles:
        try:
            if st.is_visible():
                print(f"  '{st.inner_text().strip()}'")
        except Exception:
            pass

    print("\n=== Menu items (el-menu-item) ===")
    menu_items = page.locator(".el-menu-item").all()
    for mi in menu_items:
        try:
            if mi.is_visible():
                print(f"  '{mi.inner_text().strip()}'")
        except Exception:
            pass

    # Step 4: Try clicking Template List
    tl_item = page.locator(".el-menu-item").filter(has_text="Template List")
    if tl_item.count() > 0 and tl_item.first.is_visible():
        tl_item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        print(f"\nAfter Template List click URL: {page.url}")
    else:
        # Maybe Templates has sub-menu tab
        tab = page.locator("div.el-sub-menu__title").filter(has_text="Template")
        if tab.count() > 0 and tab.first.is_visible():
            tab.first.click()
            page.wait_for_timeout(400)
            tl_item2 = page.locator(".el-menu-item").filter(has_text="Template List")
            if tl_item2.count() > 0 and tl_item2.first.is_visible():
                tl_item2.first.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(500)
                print(f"\nAfter Template List click (via sub-menu) URL: {page.url}")

    # Step 5: Check page content
    print("\n=== Visible buttons on page ===")
    btns = page.locator("button").all()
    for btn in btns:
        try:
            if btn.is_visible():
                print(f"  button: '{btn.inner_text().strip()}'")
        except Exception:
            pass

    print("\n=== Tabs on page ===")
    tabs = page.locator(".el-tabs__item").all()
    for tab in tabs:
        try:
            if tab.is_visible():
                print(f"  tab: '{tab.inner_text().strip()}'")
        except Exception:
            pass

    assert True
