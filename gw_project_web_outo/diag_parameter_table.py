"""
Diagnostic: explore the template edit page to find Parameter Table section.
Run: python -m pytest diag_parameter_table.py -v -s
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


def test_diag_parameter_table(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Go to Template List and click first custom template's yellow edit button
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Enter edit page via yellow button on last tbody (Customized)
    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0, "No custom templates found"

    first_row = rows[0]
    first_row.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    print(f"\n=== URL after clicking edit: {page.url}")

    # Print ALL buttons on the edit page
    print("\n=== All buttons on edit page ===")
    for btn in page.get_by_role("button").all():
        try:
            txt = btn.inner_text().strip()
            if txt:
                print(f"  Button: '{txt}'")
        except Exception:
            pass

    # Print ALL table sections (thead + first few rows)
    print("\n=== Tables (thead) on edit page ===")
    for i, th in enumerate(page.locator("thead").all()):
        try:
            print(f"  Table[{i}] headers: {th.inner_text().strip()[:200]}")
        except Exception:
            pass

    # Print ALL .el-tabs or tab headers
    print("\n=== Tab headers on edit page ===")
    for tab in page.locator(".el-tabs__item, .el-tab-pane").all():
        try:
            txt = tab.inner_text().strip()
            if txt:
                print(f"  Tab: '{txt[:100]}'")
        except Exception:
            pass

    # Print section/card titles
    print("\n=== Section titles (.el-card__header, h3, h4) ===")
    for h in page.locator(".el-card__header, h2, h3, h4").all():
        try:
            txt = h.inner_text().strip()
            if txt:
                print(f"  '{txt[:100]}'")
        except Exception:
            pass

    # Check for Parameter Table specific elements
    print("\n=== Checking for 'Parameter' related elements ===")
    param_count = page.locator("*").filter(has_text="Parameter").count()
    print(f"  Elements with 'Parameter' text: {param_count}")
    for el in page.locator("th, td, label, .el-card__header").all():
        try:
            txt = el.inner_text().strip()
            if "Parameter" in txt or "Multiplier" in txt or "Byte Order" in txt or "Format" in txt:
                tag = el.evaluate("el => el.tagName")
                print(f"  [{tag}] '{txt[:100]}'")
        except Exception:
            pass

    # Block Table: check if rows have edit buttons
    print("\n=== Block Table rows - checking for edit buttons ===")
    tbodies = page.locator("tbody").all()
    print(f"  Total tbodies: {len(tbodies)}")
    for i, tbody in enumerate(tbodies):
        rows2 = tbody.locator("tr").all()
        print(f"  tbody[{i}]: {len(rows2)} rows")
        if rows2:
            first = rows2[0]
            btns = first.locator("button, .el-button").all()
            for b in btns:
                try:
                    txt = b.inner_text().strip()
                    cls = b.get_attribute("class") or ""
                    print(f"    Row[0] button: '{txt}' class='{cls[:60]}'")
                except Exception:
                    pass

    assert True
