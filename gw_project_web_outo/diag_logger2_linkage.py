"""Diagnostic: get actual Log File Length → Log Interval linkage for Data Loggers 2."""
import pytest
from pages.login_page import LoginPage


def _nav_to_data_loggers2(page):
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
    item = page.locator(".el-menu-item").filter(has_text="Data Loggers 2")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


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


def test_diag_logger2_linkage(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_loggers2(page)
    page.wait_for_timeout(800)

    # Enable Logger 2
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(600)

    # Get all File Length options
    lfl_select = page.locator(".el-form-item").filter(has_text="Log File Length").first.locator(".el-select")
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

    print("\n=== Logger2 Actual Linkage ===")
    for fl, intervals in linkage.items():
        print(f"  '{fl}': {intervals}")

    assert True
