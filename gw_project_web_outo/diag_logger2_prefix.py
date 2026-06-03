"""Diagnostic: check prefix field maxlength, actual value after fill, and save result."""
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


def test_diag_prefix(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_loggers2(page)
    page.wait_for_timeout(800)

    # Enable
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(600)

    # Post Channel → Channel3
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Post Channel").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Channel3")

    # Log File Format → Json
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Format").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Json")

    # Log File Length → 10 minutes
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Length").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "10 minutes")

    # Log File Name Format → Time interval Format
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Name Format").first.locator(
        ".el-radio"
    ).filter(has_text="Time interval Format").click()
    page.wait_for_timeout(300)

    # Check prefix field
    prefix_fi = page.locator(".el-form-item").filter(has_text="Log File Name Prefix").first
    prefix_input = prefix_fi.locator("input")
    print(f"\n=== Prefix field info ===")
    print(f"  visible: {prefix_input.is_visible()}")
    print(f"  enabled: {prefix_input.is_enabled()}")
    maxlen = prefix_input.get_attribute("maxlength")
    print(f"  maxlength: {maxlen}")

    # Fill 12 chars
    prefix_input.fill("meter2_12345")
    page.wait_for_timeout(300)
    actual_val = prefix_input.input_value()
    print(f"  value after fill('meter2_12345'): '{actual_val}' (len={len(actual_val)})")

    # Log Interval → 1 minute
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log Interval").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "1 minute")

    # Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(2000)

    # Check all error indicators
    form_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    msg_warning = page.locator(".el-message--warning").count()
    msg_success = page.locator(".el-message--success").count()
    msg_any = page.locator(".el-message").count()
    print(f"\n=== After Save ===")
    print(f"  .el-form-item__error count: {form_errors}")
    print(f"  .el-message--error count: {msg_errors}")
    print(f"  .el-message--warning count: {msg_warning}")
    print(f"  .el-message--success count: {msg_success}")
    print(f"  .el-message (any) count: {msg_any}")

    # Print any message texts
    for msg in page.locator(".el-message").all():
        try:
            print(f"  message text: '{msg.inner_text().strip()}'")
        except Exception:
            pass

    # Print any form error texts
    for err in page.locator(".el-form-item__error").all():
        try:
            print(f"  form error text: '{err.inner_text().strip()}'")
        except Exception:
            pass

    assert True
