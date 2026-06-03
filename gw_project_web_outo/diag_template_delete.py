"""Diagnostic: inspect delete confirmation dialog for template."""
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


def test_diag_template_delete(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Create a template to delete
    _nav_to_templates(page)
    page.wait_for_timeout(500)

    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    template_name = f"DelTpl_{ts}"
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
    print(f"=== Template created ===")

    # Navigate to Template List
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Find the template row and click delete button
    row = page.locator("tbody tr").filter(has_text=template_name).first
    delete_btn = row.locator(".el-button--danger").first
    print(f"\n=== Clicking delete button ===")
    delete_btn.click()
    page.wait_for_timeout(1000)

    # Inspect what appeared
    print(f"\n=== Dialogs after delete click ===")
    print(f"  .el-dialog count: {page.locator('.el-dialog').count()}")
    print(f"  .el-message-box count: {page.locator('.el-message-box').count()}")
    print(f"  .el-message-box--plain count: {page.locator('.el-message-box--plain').count()}")

    # Print all visible overlays
    overlays = page.locator(".el-overlay, .v-modal, .el-overlay-message-box").all()
    print(f"  overlay count: {len(overlays)}")

    # Print all dialogs and message boxes
    for sel in [".el-message-box", ".el-dialog", "[role='dialog']"]:
        els = page.locator(sel).all()
        for el in els:
            try:
                if el.is_visible():
                    txt = el.inner_text().strip()[:200]
                    html = el.evaluate("el => el.outerHTML")[:500]
                    print(f"\n  [{sel}] text: {txt}")
                    print(f"  [{sel}] html: {html}")
            except Exception as e:
                print(f"  [{sel}] ERROR: {e}")

    # Print all buttons currently visible
    print(f"\n=== All visible buttons after delete click ===")
    btns = page.locator("button").all()
    for i, btn in enumerate(btns):
        try:
            if btn.is_visible():
                txt = btn.inner_text().strip()
                cls = btn.get_attribute("class") or ""
                print(f"  [{i}] text='{txt}' class={cls[:80]}")
        except Exception:
            pass

    assert True
