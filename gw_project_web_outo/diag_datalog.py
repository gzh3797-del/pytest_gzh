"""Diagnostic: verify actual Log File Length → Log Interval linkage."""
import pytest
from pages.login_page import LoginPage


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


def test_diag_linkage(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    page.locator(".left-nav-item").filter(has_text="Data Log").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # Enable Logger 1
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0:
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

    li_select = page.locator(".el-form-item").filter(has_text="Log Interval").first.locator(".el-select")

    linkage = {}
    for fl in file_length_options:
        # Select file length
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

        # Get interval options
        li_select.click()
        page.wait_for_timeout(400)
        interval_opts = _get_visible_options(page)
        linkage[fl] = interval_opts
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

    print("\n=== Actual Linkage (File Length → Interval options) ===")
    for fl, intervals in linkage.items():
        print(f"  '{fl}': {intervals}")

    assert True
