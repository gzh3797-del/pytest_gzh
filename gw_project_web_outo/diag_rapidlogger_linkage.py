"""Diagnostic: get actual Log File Length → Log Interval linkage for Rapid Logger."""
import pytest
from pages.login_page import LoginPage


def _nav_to_rapid_logger(page):
    if "/#/dataLog" not in page.url:
        try:
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        except Exception:
            pass
        page.locator(".left-nav-item").filter(has_text="Data Log").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    tab = page.locator("div.el-sub-menu__title").filter(has_text="Data Loggers")
    if tab.count() > 0 and tab.first.is_visible():
        tab.first.click()
        page.wait_for_timeout(400)
    item = page.locator(".el-menu-item").filter(has_text="Rapid Logger")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        print(f"\nNavigated to: {page.url}")
    else:
        print("\nWARNING: 'Rapid Logger' menu item not found or not visible")
        # Print all visible menu items
        all_items = page.locator(".el-menu-item").all()
        for mi in all_items:
            try:
                if mi.is_visible():
                    print(f"  visible menu item: '{mi.inner_text().strip()}'")
            except Exception:
                pass


def _get_visible_options(page) -> list:
    all_items = page.locator(".el-select-dropdown__item").all()
    result = []
    for item in all_items:
        try:
            if item.is_visible():
                txt = item.inner_text().strip()
                if txt:
                    result.append(txt)
        except Exception:
            pass
    return result


def test_diag_rapid_linkage(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_rapid_logger(page)
    page.wait_for_timeout(800)

    # Print all form items to understand page structure
    print("\n=== Form items on Rapid Logger page ===")
    form_items = page.locator(".el-form-item").all()
    for fi in form_items:
        try:
            label = fi.locator(".el-form-item__label").first
            label_txt = label.inner_text().strip() if label.count() > 0 else "(no label)"
            has_select = fi.locator(".el-select").count() > 0
            has_radio = fi.locator(".el-radio").count() > 0
            types = []
            if has_select:
                types.append("select")
            if has_radio:
                types.append("radio")
            print(f"  '{label_txt}': {types if types else ['other']}")
        except Exception as e:
            print(f"  ERROR: {e}")

    # Enable
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(600)

    # Get all File Length options
    lfl_fi = page.locator(".el-form-item").filter(has_text="Log File Length").first
    if lfl_fi.count() == 0:
        print("\nERROR: 'Log File Length' form item not found")
        assert True
        return

    lfl_select = lfl_fi.locator(".el-select")
    lfl_select.click()
    page.wait_for_timeout(500)
    file_length_options = _get_visible_options(page)
    print(f"\nFile Length options: {file_length_options}")
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    linkage = {}
    for fl in file_length_options:
        lfl_select.click()
        page.wait_for_timeout(400)
        items = page.locator(".el-select-dropdown__item").all()
        for item in items:
            try:
                if item.is_visible() and item.inner_text().strip() == fl:
                    item.click()
                    break
            except Exception:
                pass
        page.wait_for_timeout(400)

        li_select = page.locator(".el-form-item").filter(has_text="Log Interval").first.locator(".el-select")
        li_select.click()
        page.wait_for_timeout(400)
        interval_opts = _get_visible_options(page)
        linkage[fl] = interval_opts
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

    print("\n=== RapidLogger Actual Linkage ===")
    for fl, intervals in linkage.items():
        print(f"  '{fl}': {intervals}")

    assert True
