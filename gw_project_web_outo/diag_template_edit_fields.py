"""Diagnostic: inspect field states in template edit page."""
import time
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


def _click_visible_option(page, option_text: str):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    return False


def test_diag_template_edit_fields(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Create a template first
    _nav_to_templates(page)
    page.wait_for_timeout(500)

    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    template_name = f"DiagTpl_{ts}"
    print(f"\n=== Creating template: {template_name} ===")

    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(template_name)
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("0001")
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)

    # Go to Template List and click edit
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    row = page.locator("tbody tr").filter(has_text=template_name).first
    row.locator(".el-button--primary").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    print(f"\n=== Edit page URL: {page.url} ===")

    # Inspect each form item
    field_labels = ["Template Name", "Version", "Typical Model", "Wiring Configuration", "Function", "Start", "Count"]
    for label in field_labels:
        fi = page.locator(".el-form-item").filter(has_text=label).first
        if fi.count() == 0:
            print(f"  [{label}] NOT FOUND")
            continue

        # Check input
        inp = fi.locator("input").first
        if inp.count() > 0:
            disabled_attr = inp.get_attribute("disabled")
            readonly_attr = inp.get_attribute("readonly")
            el_input_cls = ""
            try:
                el_input_cls = fi.locator(".el-input").first.get_attribute("class") or ""
            except Exception:
                pass
            fi_cls = fi.get_attribute("class") or ""
            print(f"  [{label}] input: disabled={disabled_attr!r} readonly={readonly_attr!r} "
                  f"el-input-cls={el_input_cls[:80]} fi-cls={fi_cls[:80]}")
            # Try full form item HTML
            try:
                html = fi.evaluate("el => el.outerHTML")[:400]
                print(f"  [{label}] html: {html}")
            except Exception as e:
                print(f"  [{label}] html error: {e}")
        else:
            # Check select
            sel = fi.locator(".el-select").first
            if sel.count() > 0:
                sel_cls = sel.get_attribute("class") or ""
                print(f"  [{label}] select-cls={sel_cls[:80]}")
            else:
                print(f"  [{label}] no input or select found")

    assert True
