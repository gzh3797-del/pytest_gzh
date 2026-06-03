"""
Diagnostic: find where Ack Status column lives and how to verify it.
Check: device Alarm Logs, global Alarm Logs, and Add Alarm form toggle.
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

        # ── 1. Check global Alarm page structure ──────────────────────────────
        print("=== Navigate to global Alarm (left nav) ===")
        page.locator(".left-nav-item").filter(has_text="Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("URL:", page.url)

        print("\n=== Global Alarm sub-menus ===")
        for el in page.locator(".left-nav-item, [role='menuitem'], .el-menu-item").all():
            try:
                txt = el.inner_text().strip()
                if txt and len(txt) < 50:
                    print(f"  '{txt}'")
            except Exception:
                pass

        # Check for Alarm Logs tab
        print("\n=== Looking for Alarm Logs link/tab on global Alarm page ===")
        for sel in ["a", ".el-tabs__item", "[role='tab']", "li"]:
            els = page.locator(sel).filter(has_text="Alarm Logs")
            if els.count() > 0:
                print(f"  Found via '{sel}': count={els.count()}")
                for i in range(els.count()):
                    print(f"    [{i}] tag={els.nth(i).evaluate('e=>e.tagName')} "
                          f"class='{els.nth(i).get_attribute('class')}'")
                break

        # Try clicking Alarm Logs on global page
        alarm_logs_link = page.locator(".el-tabs__item, a, li").filter(has_text="Alarm Logs").first
        if alarm_logs_link.count() > 0:
            alarm_logs_link.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(800)
            print("\nURL after clicking Alarm Logs:", page.url)

            print("\n=== Global Alarm Logs table columns ===")
            for el in page.locator("th").all():
                try:
                    txt = el.inner_text().strip()
                    if txt:
                        print(f"  '{txt}'")
                except Exception:
                    pass

            ack = page.locator("th").filter(has_text="Ack Status")
            print(f"\n'Ack Status' in global Alarm Logs: {ack.count() > 0}")
        else:
            print("  No Alarm Logs link found on global Alarm page")

        # ── 2. Check device Alarm Config Add Alarm form (toggle) ─────────────
        print("\n\n=== Navigate to Acu4100 Alarm Config ===")
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
        print("Alarm Config URL:", page.url)

        # Check for any existing alarm_4100 with Ack column
        print("\n=== Alarm Config list table columns ===")
        for el in page.locator("th").all():
            try:
                txt = el.inner_text().strip()
                if txt:
                    print(f"  '{txt}'")
            except Exception:
                pass

        # Open Add Alarm form and inspect toggle
        page.get_by_role("button", name="Add Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)

        print("\n=== Add Alarm form elements (toggles/switches) ===")
        for el in page.locator(".el-switch, [role='switch'], input[type='checkbox']").all():
            try:
                cls = el.get_attribute("class") or ""
                aria = el.get_attribute("aria-checked") or ""
                parent = el.evaluate("e => e.parentElement.innerText").strip()[:60]
                print(f"  class='{cls}' aria-checked='{aria}' parent='{parent}'")
            except Exception as ex:
                print(f"  error: {ex}")

        print("\n=== Add Alarm form labels ===")
        for el in page.locator("label, .el-form-item__label").all():
            try:
                print(f"  '{el.inner_text().strip()}'")
            except Exception:
                pass

        # ── 3. Check device Alarm Logs columns after alarm_4100 exists ────────
        print("\n\n=== Device Alarm Logs page columns ===")
        alarm_sub = page.locator("li.el-sub-menu").filter(has_text="Alarm")
        alarm_sub.click()
        page.wait_for_timeout(300)
        page.get_by_role("menuitem", name="Alarm Logs").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("Alarm Logs URL:", page.url)

        print("Table columns:")
        for el in page.locator("th").all():
            try:
                txt = el.inner_text().strip()
                if txt:
                    print(f"  '{txt}'")
            except Exception:
                pass

        print(f"\n'Ack Status' in device Alarm Logs: {page.locator('th').filter(has_text='Ack Status').count() > 0}")
        print(f"alarm_4100 rows in device Alarm Logs: {page.locator('tr.el-table__row').filter(has_text='alarm_4100').count()}")

        # ── 4. Check first few rows for Ack-related text ──────────────────────
        print("\n=== First 3 rows content ===")
        rows = page.locator("tr.el-table__row")
        for i in range(min(rows.count(), 3)):
            print(f"  Row {i}: '{rows.nth(i).inner_text().strip()[:200]}'")

        page.wait_for_timeout(1000)
        browser.close()

if __name__ == "__main__":
    run()
