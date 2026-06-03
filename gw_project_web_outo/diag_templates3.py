"""Diagnostic: click the green button on Template List and inspect the dialog."""
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


def test_diag_templates3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_template_list(page)
    page.wait_for_timeout(500)

    # Print first row content to understand table structure
    print("\n=== First row HTML ===")
    rows = page.locator("tbody tr").all()
    if rows:
        html = rows[0].evaluate("el => el.outerHTML")
        print(html[:600])

    # Check table header
    print("\n=== Table headers ===")
    headers = page.locator("thead th").all()
    for th in headers:
        try:
            print(f"  '{th.inner_text().strip()}'")
        except Exception:
            pass

    # Click the green (success) button — likely "Add Template"
    green_btn = page.locator(".el-button--success").first
    print(f"\n=== Green button visible: {green_btn.is_visible()} ===")
    green_btn.click()
    page.wait_for_timeout(800)

    # Check if dialog appeared
    dialog = page.locator(".el-dialog")
    print(f"\n=== Dialog count: {dialog.count()} ===")
    if dialog.count() > 0:
        print(f"Dialog visible: {dialog.first.is_visible()}")
        print(f"Dialog HTML (first 1000 chars): {dialog.first.evaluate('el => el.outerHTML')[:1000]}")

        # Print form items in dialog
        print("\n=== Dialog form items ===")
        form_items = dialog.first.locator(".el-form-item").all()
        for fi in form_items:
            try:
                label = fi.locator(".el-form-item__label").first.inner_text().strip()
                has_select = fi.locator(".el-select").count() > 0
                has_input = fi.locator("input").count() > 0
                has_radio = fi.locator(".el-radio").count() > 0
                types = []
                if has_select: types.append("select")
                if has_input and not has_select: types.append("input")
                if has_radio: types.append("radio")
                print(f"  '{label}': {types if types else ['other']}")
            except Exception as e:
                print(f"  ERROR: {e}")

        # Print dialog buttons
        print("\n=== Dialog buttons ===")
        dlg_btns = dialog.first.locator("button").all()
        for btn in dlg_btns:
            try:
                if btn.is_visible():
                    txt = btn.inner_text().strip()
                    html = btn.evaluate("el => el.outerHTML")[:200]
                    print(f"  text='{txt}' html={html}")
            except Exception:
                pass

        # Check for tabs in dialog
        print("\n=== Dialog tabs ===")
        tabs = dialog.first.locator(".el-tabs__item").all()
        for tab in tabs:
            try:
                print(f"  tab: '{tab.inner_text().strip()}' visible={tab.is_visible()}")
            except Exception:
                pass

        # Close dialog
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

    # Also check URL navigation — did a new page open?
    print(f"\nURL after green button click: {page.url}")

    # Check if page changed to a new template creation page
    print("\n=== Page buttons after green click ===")
    all_btns = page.locator(".el-button").all()
    for btn in all_btns:
        try:
            if btn.is_visible():
                txt = btn.inner_text().strip()
                cls = btn.get_attribute("class") or ""
                print(f"  text='{txt}' class={cls[:80]}")
        except Exception:
            pass

    assert True
