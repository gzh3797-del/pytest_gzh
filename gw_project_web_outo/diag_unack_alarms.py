"""
Diagnostic: inspect Unacknowledged Alarms page and the Monitor Label column.
Check if alarm_4100 appears and what the page structure is.
"""
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, HEADLESS, SLOW_MO
from pages.login_page import LoginPage

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=100)
        ctx = browser.new_context(base_url=BASE_URL, ignore_https_errors=True,
                                  viewport={"width": 1280, "height": 720})
        page = ctx.new_page()
        lp = LoginPage(page)
        lp.open()
        lp.login()

        # Navigate to global Alarm > Unacknowledged Alarms
        page.locator(".left-nav-item").filter(has_text="Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)
        print("URL:", page.url)

        print("\n=== Menu items on Alarm page ===")
        for el in page.locator(".el-menu-item").all():
            try:
                print(f"  '{el.inner_text().strip()}'")
            except Exception:
                pass

        # Check what page we're on - should be activeAlarm
        print(f"On activeAlarm page: {'activeAlarm' in page.url}")

        print("\n=== Unacknowledged Alarms table columns ===")
        for el in page.locator("th").all():
            try:
                txt = el.inner_text().strip()
                if txt:
                    print(f"  '{txt}'")
            except Exception:
                pass

        print("\n=== All table rows content ===")
        rows = page.locator("tr")
        print(f"Total rows (including header): {rows.count()}")
        for i in range(rows.count()):
            row = rows.nth(i)
            try:
                txt = row.inner_text().strip()
                if txt and txt != "" and not txt.startswith("No"):
                    print(f"  row[{i}]: '{txt[:150]}'")
            except Exception:
                pass

        print(f"\nalarm_4100 visible: {page.locator('tr').filter(has_text='alarm_4100').count() > 0}")
        print(f"Total alarm rows: {page.locator('tr.el-table__row').count()}")

        # Print each alarm row
        alarm_rows = page.locator("tr.el-table__row")
        for i in range(alarm_rows.count()):
            row = alarm_rows.nth(i)
            try:
                txt = row.inner_text().strip().replace('\n', ' | ')
                print(f"  Alarm row[{i}]: '{txt[:200]}'")
            except Exception:
                pass

        # Check Monitor Label column
        print("\n=== Monitor Label column values ===")
        for row in page.locator("tr.el-table__row").all():
            cells = row.locator("td")
            for i in range(cells.count()):
                txt = cells.nth(i).inner_text().strip()
                if txt == "alarm_4100":
                    print(f"  Found 'alarm_4100' in cell {i} of row")

        page.wait_for_timeout(2000)
        browser.close()

if __name__ == "__main__":
    run()
