"""Diagnostic: inspect Add Alarm form selectors."""
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

        # Navigate to Physical Devices
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        # Click Acu4100
        page.locator("tr.el-table__row").filter(has_text=_DEVICE_NAME).first.locator("td").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)
        print("Device page URL:", page.url)

        # Click Alarm Config (must first expand Alarm submenu)
        alarm_sub = page.locator("li.el-sub-menu").filter(has_text="Alarm")
        alarm_sub.click()
        page.wait_for_timeout(400)
        page.get_by_role("menuitem", name="Alarm Config").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)
        print("Alarm Config URL:", page.url)

        # Click Add Alarm
        page.get_by_role("button", name="Add Alarm").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(600)
        print("Add Alarm URL:", page.url)

        # Inspect form fields
        print("\n=== Form input placeholders ===")
        for inp in page.locator("input, textarea").all():
            try:
                ph = inp.get_attribute("placeholder") or ""
                if ph:
                    print(f"  placeholder='{ph}'")
            except Exception:
                pass

        print("\n=== el-select elements ===")
        for i, sel in enumerate(page.locator(".el-select").all()):
            try:
                txt = sel.inner_text().strip()[:50]
                cls = sel.get_attribute("class") or ""
                print(f"  [{i}] text='{txt}' class='{cls}'")
            except Exception:
                pass

        # Click the Parameter dropdown
        print("\n=== Clicking Parameter select ===")
        param_item = page.locator(".el-form-item").filter(has_text="Parameter")
        print(f"  .el-form-item with 'Parameter': count={param_item.count()}")

        sel_in_param = param_item.locator(".el-select")
        print(f"  .el-select inside it: count={sel_in_param.count()}")
        sel_in_param.first.click()
        page.wait_for_timeout(600)

        print("\n=== Options after clicking Parameter ===")
        # Try various option selectors
        for sel in [
            "li.el-select-dropdown__item",
            "[role='option']",
            ".el-select-dropdown__item",
            ".el-dropdown-menu__item",
            "li",
        ]:
            els = page.locator(sel)
            if els.count() > 0:
                print(f"  Selector '{sel}': count={els.count()}")
                for i in range(min(els.count(), 5)):
                    try:
                        print(f"    [{i}] '{els.nth(i).inner_text().strip()}'")
                    except Exception:
                        pass
                break

        # Type to filter
        print("\n=== Type 'System' to filter ===")
        param_input = param_item.locator("input")
        param_input.fill("System")
        page.wait_for_timeout(600)

        for sel in [
            "li.el-select-dropdown__item",
            "[role='option']",
            ".el-select-dropdown__item",
        ]:
            els = page.locator(sel)
            if els.count() > 0:
                print(f"  After typing - Selector '{sel}': count={els.count()}")
                for i in range(min(els.count(), 10)):
                    try:
                        print(f"    [{i}] '{els.nth(i).inner_text().strip()}'")
                    except Exception:
                        pass
                break

        page.wait_for_timeout(2000)
        browser.close()

if __name__ == "__main__":
    run()
