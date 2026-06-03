"""Diagnostic: explore Edit Parameter dialog options and error messages."""
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


def _click_visible_option(page, option_text: str = ""):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(300)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def _list_options(page, label):
    """Open a dropdown by label and list all options."""
    print(f"\n=== Options for '{label}' dropdown ===")
    options = []
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible():
                txt = item.inner_text().strip()
                options.append(txt)
                print(f"  '{txt}'")
        except Exception:
            pass
    if not options:
        print("  (none found)")
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return options


def test_diag_param_edit_options(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="Template List").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    last_tbody = page.locator("tbody").last
    rows = last_tbody.locator("tr").all()
    assert len(rows) > 0
    last_tbody.locator("tr").first.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    tbodies = page.locator("tbody").all()
    target_tbody = None
    for i, tb in enumerate(tbodies):
        if i > 0 and tb.locator("tr").count() > 0:
            target_tbody = tb
            break

    assert target_tbody is not None
    target_tbody.locator("tr").first.locator(".el-button--primary").first.click()
    page.wait_for_timeout(500)

    dialog = page.locator(".el-dialog").filter(has_text="Edit Parameter").first
    assert dialog.count() > 0, "Dialog not found"

    # List Block options
    dialog.locator(".el-form-item").filter(has_text="Block").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    block_opts = _list_options(page, "Block")

    # List Address Format options
    dialog.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _list_options(page, "Address Format")

    # Select Hex in Address Format
    dialog.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "Hex")

    # List Data Format options
    dialog.locator(".el-form-item").filter(has_text="Data Format").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _list_options(page, "Data Format")

    # Select first Data Format
    dialog.locator(".el-form-item").filter(has_text="Data Format").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    # List Byte Order options
    dialog.locator(".el-form-item").filter(has_text="Byte Order").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _list_options(page, "Byte Order")

    # Select first Byte Order
    dialog.locator(".el-form-item").filter(has_text="Byte Order").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    # Select Block (first option)
    dialog.locator(".el-form-item").filter(has_text="Block").first.locator(".el-select").first.click()
    page.wait_for_timeout(300)
    _click_visible_option(page, "")

    # Fill Address = 0001 (small valid hex)
    addr_fi = dialog.locator(".el-form-item").filter(has_text="Address").filter(has_not_text="Format").first
    addr_inp = addr_fi.locator("input").first
    addr_inp.click()
    addr_inp.fill("0001")
    page.wait_for_timeout(200)

    # Fill Multiplier = 1
    mul_fi = dialog.locator(".el-form-item").filter(has_text="Multiplier").first
    mul_inp = mul_fi.locator("input").first
    mul_inp.click()
    mul_inp.fill("1")
    page.wait_for_timeout(200)

    # Print current dialog state before save
    print("\n=== Dialog form state before Save ===")
    for fi in dialog.locator(".el-form-item").all():
        try:
            if fi.is_visible():
                label = fi.locator("label").first.inner_text().strip()
                val = ""
                if fi.locator("input").count() > 0:
                    val = fi.locator("input").first.input_value()
                elif fi.locator(".el-select__wrapper").count() > 0:
                    val = fi.locator(".el-select__wrapper").first.inner_text().strip()
                print(f"  '{label}': '{val}'")
        except Exception:
            pass

    # Click Save
    dialog.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # Check errors
    print("\n=== After Save ===")
    errors = page.locator(".el-form-item__error").all()
    for e in errors:
        try:
            if e.is_visible():
                print(f"  Form error: '{e.inner_text()}'")
        except Exception:
            pass
    msgs = page.locator(".el-message").all()
    for m in msgs:
        try:
            if m.is_visible():
                print(f"  Message: '{m.inner_text()}'")
        except Exception:
            pass
    print(f"  Dialog still visible: {dialog.is_visible()}")

    assert True
