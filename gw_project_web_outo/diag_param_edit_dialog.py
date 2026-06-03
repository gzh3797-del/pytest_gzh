"""
Diagnostic: click blue Action button in Parameter Table row and explore edit dialog.
Run: python -m pytest diag_param_edit_dialog.py -v -s
"""
from pages.login_page import LoginPage


def _nav_to_templates(page):
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


def test_diag_param_edit_dialog(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Enter template edit page via yellow button
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0, "No custom templates"
    last_tbody.locator("tr").first.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Find Parameter Table (tbody[2]) and click blue button on first row with data
    tbodies = page.locator("tbody").all()
    print(f"\nTotal tbodies: {len(tbodies)}")

    target_tbody = None
    for i, tb in enumerate(tbodies):
        row_count = tb.locator("tr").count()
        print(f"  tbody[{i}]: {row_count} rows")
        if row_count > 0 and i > 0:  # Skip Block Table (i=0), find first Parameter Table with rows
            target_tbody = tb
            print(f"  --> Using tbody[{i}] as Parameter Table")
            break

    assert target_tbody is not None, "No Parameter Table with rows found"

    first_param_row = target_tbody.locator("tr").first
    row_text = first_param_row.inner_text().strip()[:200]
    print(f"\nFirst Parameter Table row text: {row_text}")

    # Click the blue Action button in the first param row
    blue_btn = first_param_row.locator(".el-button--primary").first
    print(f"\nBlue button count in first row: {blue_btn.count()}")
    if blue_btn.count() == 0:
        # Try any button
        all_btns = first_param_row.locator("button, .el-button").all()
        print(f"All buttons in first row: {len(all_btns)}")
        for b in all_btns:
            try:
                print(f"  btn: '{b.inner_text().strip()}' class='{b.get_attribute('class')}'")
            except Exception:
                pass

    blue_btn.first.click()
    page.wait_for_timeout(800)

    print(f"\n=== After clicking blue button ===")
    print(f"URL: {page.url}")

    # Check if a dialog/drawer appeared
    dialog = page.locator(".el-dialog, .el-drawer")
    print(f"Dialogs/Drawers visible: {dialog.count()}")
    for d in dialog.all():
        try:
            if d.is_visible():
                print(f"  Dialog text (first 300): {d.inner_text()[:300]}")
        except Exception:
            pass

    # Check form items
    print("\n=== Form items in dialog ===")
    for fi in page.locator(".el-form-item").all():
        try:
            if fi.is_visible():
                label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
                # Check for input, select
                has_input = fi.locator("input").count() > 0
                has_select = fi.locator(".el-select").count() > 0
                print(f"  Form item: label='{label}' has_input={has_input} has_select={has_select}")
        except Exception:
            pass

    # Check visible buttons after click
    print("\n=== Visible buttons after click ===")
    for btn in page.get_by_role("button").all():
        try:
            if btn.is_visible():
                txt = btn.inner_text().strip()
                if txt:
                    print(f"  Button: '{txt}'")
        except Exception:
            pass

    # Check visible select dropdowns
    print("\n=== Visible selects ===")
    for sel in page.locator(".el-select").all():
        try:
            if sel.is_visible():
                val = sel.locator(".el-select__wrapper, .el-input__inner").first.inner_text().strip()
                print(f"  Select value: '{val}'")
        except Exception:
            pass

    assert True
