"""Diagnostic: print all form item labels and their input types on Data Loggers 2 page."""
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


def _click_option(page, option_text: str):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and item.inner_text().strip() == option_text:
                item.click()
                page.wait_for_timeout(400)
                return
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)


def test_diag_logger2(login_page: LoginPage):
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

    # Print all form item labels after enable
    print("\n=== Form items after Enable ===")
    form_items = page.locator(".el-form-item").all()
    for fi in form_items:
        try:
            label = fi.locator(".el-form-item__label").first
            label_txt = label.inner_text().strip() if label.count() > 0 else "(no label)"
            has_select = fi.locator(".el-select").count() > 0
            has_input = fi.locator("input").count() > 0
            has_radio = fi.locator(".el-radio").count() > 0
            has_checkbox = fi.locator(".el-checkbox").count() > 0
            types = []
            if has_select:
                types.append("select")
            if has_input and not has_select:
                types.append("input")
            if has_radio:
                types.append("radio")
            if has_checkbox:
                types.append("checkbox")
            print(f"  '{label_txt}': {types if types else ['(no input)']}")
        except Exception as e:
            print(f"  ERROR: {e}")

    # Select Post Channel → Channel3
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Post Channel").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Channel3")

    # Select Log File Format → Json
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Format").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Json")

    # Print form items again after selecting Json
    print("\n=== Form items after Json selected ===")
    form_items = page.locator(".el-form-item").all()
    for fi in form_items:
        try:
            label = fi.locator(".el-form-item__label").first
            label_txt = label.inner_text().strip() if label.count() > 0 else "(no label)"
            has_select = fi.locator(".el-select").count() > 0
            has_input = fi.locator("input").count() > 0
            has_radio = fi.locator(".el-radio").count() > 0
            types = []
            if has_select:
                types.append("select")
            if has_input and not has_select:
                types.append("input")
            if has_radio:
                types.append("radio")
            print(f"  '{label_txt}': {types if types else ['(no input)']}")
        except Exception as e:
            print(f"  ERROR: {e}")

    assert True
