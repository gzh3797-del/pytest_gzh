"""
Diagnostic: inspect Acu4100 device Alarm tab navigation selectors
and Alarm Logs table column structure.
"""
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, HEADLESS, SLOW_MO
from pages.login_page import LoginPage

DEVICE_NAME = "Acu4100"

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        ctx = browser.new_context(base_url=BASE_URL, ignore_https_errors=True,
                                  viewport={"width": 1280, "height": 720})
        page = ctx.new_page()
        lp = LoginPage(page)
        lp.open()
        lp.login()

        # ── 1. Navigate to Physical Devices ──────────────────────────────────
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("URL after Physical Devices click:", page.url)

        # ── 2. Click Acu4100 link ─────────────────────────────────────────────
        print("\n=== Looking for device name element ===")
        # Try various selectors for the device name link
        for sel in [
            f"a:has-text('{DEVICE_NAME}')",
            f"td a:has-text('{DEVICE_NAME}')",
            f"[class*='device'] a:has-text('{DEVICE_NAME}')",
            f"text='{DEVICE_NAME}'",
        ]:
            els = page.locator(sel)
            if els.count() > 0:
                print(f"  Found via selector: {sel} -> count={els.count()}")
                break

        # Print all links on the page
        all_links = page.locator("a")
        print(f"  Total <a> tags: {all_links.count()}")
        for i in range(min(all_links.count(), 30)):
            el = all_links.nth(i)
            try:
                txt = el.inner_text().strip()
                if txt:
                    print(f"    link[{i}] text='{txt}' href='{el.get_attribute('href')}'")
            except Exception:
                pass

        # Also print table cell content
        print("\n=== Table cells with device-like content ===")
        tds = page.locator("td")
        for i in range(min(tds.count(), 50)):
            el = tds.nth(i)
            try:
                txt = el.inner_text().strip()
                if DEVICE_NAME in txt or "Acu" in txt:
                    print(f"  td[{i}] text='{txt}' class='{el.get_attribute('class')}'")
            except Exception:
                pass

        # Inspect inner HTML of the Acu4100 Device Name cell
        print("\n=== Inner HTML of first td with 'Acu4100' text ===")
        first_td = page.locator("td").filter(has_text=DEVICE_NAME).first
        try:
            print(first_td.inner_html())
        except Exception as ex:
            print(f"  error: {ex}")

        # Try clicking the row
        print("\n=== Try clicking tr row ===")
        rows = page.locator("tr.el-table__row").filter(has_text=DEVICE_NAME)
        print(f"  Found {rows.count()} rows")
        for i in range(rows.count()):
            row = rows.nth(i)
            try:
                first_cell = row.locator("td").first
                print(f"  row[{i}] first cell text='{first_cell.inner_text().strip()}'")
            except Exception as ex:
                print(f"  row[{i}] error: {ex}")

        # Click the first cell of the Acu4100 row
        try:
            target_row = page.locator("tr.el-table__row").filter(has_text=DEVICE_NAME).first
            first_cell = target_row.locator("td").first
            print(f"  Clicking first cell: '{first_cell.inner_text().strip()}'")
            first_cell.click()
        except Exception as ex:
            print(f"  Click failed: {ex}")
            # Try clicking the span inside
            page.locator("td").filter(has_text=DEVICE_NAME).first.locator("span").click()

        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("URL on device detail page:", page.url)

        # ── 3. Inspect the top nav bar ────────────────────────────────────────
        print("\n=== Top nav links/buttons on device detail page ===")
        for el in page.locator("nav a, nav li, .el-menu-item, .el-submenu__title").all():
            try:
                print(f"  tag={el.evaluate('e=>e.tagName')} "
                      f"class={el.get_attribute('class')} "
                      f"text='{el.inner_text().strip()}'")
            except Exception:
                pass

        # ── 4. Try clicking the Alarm tab ─────────────────────────────────────
        print("\n=== Looking for Alarm tab ===")
        # Try different selectors
        candidates = page.locator("a, li").filter(has_text="Alarm")
        print(f"  Found {candidates.count()} elements with text 'Alarm'")
        for i in range(candidates.count()):
            el = candidates.nth(i)
            try:
                print(f"  [{i}] tag={el.evaluate('e=>e.tagName')} "
                      f"class='{el.get_attribute('class')}' "
                      f"text='{el.inner_text().strip()[:40]}'")
            except Exception as ex:
                print(f"  [{i}] error: {ex}")

        # ── 5. Click the device-level Alarm tab ───────────────────────────────
        # Try clicking - look for one that's NOT the left nav item
        clicked = False
        for i in range(candidates.count()):
            el = candidates.nth(i)
            try:
                cls = el.get_attribute("class") or ""
                if "left-nav" not in cls:
                    print(f"\nClicking candidate [{i}] ...")
                    el.click()
                    page.wait_for_timeout(500)
                    print("URL after click:", page.url)
                    clicked = True
                    break
            except Exception as ex:
                print(f"  Could not click [{i}]: {ex}")

        if not clicked:
            print("Could not find/click Alarm tab")
            browser.close()
            return

        # ── 6. Check dropdown menu appeared ───────────────────────────────────
        print("\n=== Dropdown menu items after Alarm click ===")
        for el in page.locator(".el-menu--popup li, [role='menuitem'], .el-dropdown-menu li").all():
            try:
                print(f"  text='{el.inner_text().strip()}'")
            except Exception:
                pass

        # ── 7. Navigate to Alarm Config ───────────────────────────────────────
        try:
            page.get_by_role("menuitem", name="Alarm Config").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
            print("\nURL after Alarm Config click:", page.url)
        except Exception as ex:
            print(f"Could not click Alarm Config: {ex}")
            # Try generic approach
            page.locator("a, li").filter(has_text="Alarm Config").first.click()
            page.wait_for_timeout(500)
            print("URL (fallback):", page.url)

        # ── 8. Inspect Alarm Config list page ─────────────────────────────────
        print("\n=== Add Alarm button ===")
        add_btn = page.get_by_role("button", name="Add Alarm")
        print(f"  'Add Alarm' button visible: {add_btn.is_visible()}")

        # ── 9. Navigate to Alarm Logs ─────────────────────────────────────────
        print("\n=== Navigate to Alarm > Alarm Logs ===")
        # Click Alarm tab again
        for i in range(candidates.count()):
            el = candidates.nth(i)
            try:
                cls = el.get_attribute("class") or ""
                if "left-nav" not in cls:
                    el.click()
                    page.wait_for_timeout(500)
                    break
            except Exception:
                pass

        try:
            page.get_by_role("menuitem", name="Alarm Logs").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(800)
            print("URL after Alarm Logs click:", page.url)
        except Exception as ex:
            print(f"Could not click Alarm Logs: {ex}")

        # ── 10. Inspect Alarm Logs table headers ──────────────────────────────
        print("\n=== Alarm Logs table columns (th elements) ===")
        for el in page.locator("th, .el-table__header th, thead td").all():
            try:
                txt = el.inner_text().strip()
                if txt:
                    print(f"  '{txt}'")
            except Exception:
                pass

        # ── 11. Check for Ack Status column ───────────────────────────────────
        ack_col = page.locator("th").filter(has_text="Ack Status")
        print(f"\n'Ack Status' column header visible: {ack_col.count() > 0}")

        page.wait_for_timeout(2000)
        browser.close()

if __name__ == "__main__":
    run()
