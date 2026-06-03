"""
Diagnostic: compare auto_test1 (in Unacknowledged Alarms) vs alarm_4100 (not in).
Check their edit forms for differences.
Also: try adding alarm_4100 fresh and wait to see if it appears.
"""
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, HEADLESS, SLOW_MO
from pages.login_page import LoginPage

_DEVICE_NAME = "Acu4100"

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=100)
        ctx = browser.new_context(base_url=BASE_URL, ignore_https_errors=True,
                                  viewport={"width": 1280, "height": 720})
        page = ctx.new_page()
        lp = LoginPage(page)
        lp.open()
        lp.login()

        # Navigate to Acu4100 Alarm Config
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.locator("tr.el-table__row").filter(has_text=_DEVICE_NAME).first.locator("td").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)
        alarm_sub = page.locator("li.el-sub-menu").filter(has_text="Alarm")
        alarm_sub.click()
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="Alarm Config").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)

        print("=== Current Alarm Config list ===")
        rows = page.locator("tr.el-table__row")
        for i in range(rows.count()):
            txt = rows.nth(i).inner_text().strip().replace('\n', ' | ')
            print(f"  row[{i}]: '{txt[:200]}'")

        # ── 1. Open edit form for auto_test1 ──────────────────────────────────
        print("\n=== Edit form for auto_test1 ===")
        auto_row = page.locator("tr.el-table__row").filter(has_text="auto_test1").first
        if auto_row.count() > 0:
            # Click edit button (first button in row)
            auto_row.locator("button").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(600)

            # Print all form elements
            print("  Form elements:")
            for inp in page.locator("input").all():
                try:
                    ph = inp.get_attribute("placeholder") or ""
                    val = inp.input_value()
                    disabled = inp.get_attribute("disabled")
                    print(f"    input ph='{ph}' value='{val}' disabled={disabled}")
                except Exception:
                    pass

            # Check switches
            for sw in page.locator(".el-switch").all():
                try:
                    cls = sw.get_attribute("class") or ""
                    parent = sw.evaluate("e => e.closest('.el-form-item')?.innerText || e.parentElement?.innerText or ''").strip()[:60]
                    print(f"    switch class='{cls}' parent='{parent[:60]}'")
                except Exception as ex:
                    print(f"    switch error: {ex}")

            # Check all select values
            for sel in page.locator(".el-select").all():
                try:
                    txt = sel.inner_text().strip()[:40]
                    parent = sel.evaluate("e => e.closest('.el-form-item')?.querySelector('label')?.innerText || ''")
                    print(f"    select label='{parent}' value='{txt}'")
                except Exception:
                    pass

            # Go back
            page.get_by_role("button", name="Cancel").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        else:
            print("  auto_test1 not found in Alarm Config")

        # ── 2. Add alarm_4100 fresh (if not exists) ────────────────────────────
        alarm_rows = page.locator("tr.el-table__row").filter(has_text="alarm_4100")
        if alarm_rows.count() > 0:
            print("\n=== alarm_4100 already exists, opening edit form ===")
            alarm_rows.first.locator("button").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(600)

            print("  Form elements:")
            for inp in page.locator("input").all():
                try:
                    ph = inp.get_attribute("placeholder") or ""
                    val = inp.input_value()
                    disabled = inp.get_attribute("disabled")
                    print(f"    input ph='{ph}' value='{val}' disabled={disabled}")
                except Exception:
                    pass

            for sw in page.locator(".el-switch").all():
                try:
                    cls = sw.get_attribute("class") or ""
                    parent = sw.evaluate("e => e.parentElement?.innerText or ''").strip()[:60]
                    print(f"    switch class='{cls}' parent='{parent[:60]}'")
                except Exception as ex:
                    print(f"    switch error: {ex}")

            page.get_by_role("button", name="Cancel").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        else:
            print("\n=== alarm_4100 not found; will Add Alarm ===")
            page.get_by_role("button", name="Add Alarm").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)

            print("=== Add Alarm form - BEFORE any fill ===")
            for sw in page.locator(".el-switch").all():
                try:
                    cls = sw.get_attribute("class") or ""
                    aria = sw.get_attribute("aria-checked") or ""
                    parent = sw.evaluate("e => e.parentElement?.innerText or ''").strip()[:80]
                    print(f"    switch class='{cls}' aria='{aria}' parent='{parent}'")
                except Exception as ex:
                    print(f"    switch error: {ex}")

            # Click Cancel (don't save)
            page.get_by_role("button", name="Cancel").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)

        # ── 3. Check current Unacknowledged Alarms ─────────────────────────────
        print("\n=== Current Unacknowledged Alarms ===")
        page.locator(".left-nav-item").filter(has_text="Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("URL:", page.url)

        for row in page.locator("tr.el-table__row").all():
            try:
                txt = row.inner_text().strip().replace('\n', ' | ')
                print(f"  '{txt[:200]}'")
            except Exception:
                pass

        print(f"\nalarm_4100 in Unacknowledged Alarms: {page.locator('tr').filter(has_text='alarm_4100').count() > 0}")

        page.wait_for_timeout(1000)
        browser.close()

if __name__ == "__main__":
    run()
