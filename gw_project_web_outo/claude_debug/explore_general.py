"""Explore General page Session Timeout input element."""
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()

        # Login
        page.goto(BASE_URL + "/#/login")
        page.wait_for_load_state("networkidle")
        page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
        page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle")
        try:
            page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            pass

        # Navigate to User Management > General
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
        page.get_by_role("menuitem", name="General").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        print(f"URL: {page.url}")

        # Check spinbutton
        spinbtns = page.get_by_role("spinbutton").all()
        print(f"spinbutton count: {len(spinbtns)}")

        # Check all inputs
        inputs = page.locator("input").all()
        print(f"Total inputs: {len(inputs)}")
        for i, inp in enumerate(inputs):
            try:
                t = inp.get_attribute("type")
                ph = inp.get_attribute("placeholder")
                name = inp.get_attribute("name")
                aria = inp.get_attribute("aria-label")
                val = inp.input_value()
                print(f"  input[{i}]: type={t}, placeholder={ph!r}, name={name}, aria-label={aria}, value={val!r}")
            except Exception as e:
                print(f"  input[{i}]: error {e}")

        # Try filling spinbutton
        try:
            page.get_by_role("spinbutton").fill("10", timeout=3000)
            print("spinbutton fill succeeded!")
        except Exception as e:
            print(f"spinbutton fill failed: {e}")

        ctx.close()
        browser.close()


if __name__ == "__main__":
    run()
