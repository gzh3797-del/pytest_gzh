"""Explore Add User form to check checkbox defaults."""
from playwright.sync_api import sync_playwright
from config.settings import BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()

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

        # Navigate to User Configuration
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
        page.get_by_role("menuitem", name="User Configuration").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        # Open Add User dialog
        page.get_by_role("button", name="Add User").click()
        page.wait_for_timeout(1000)

        # Check all checkboxes and their default state
        checkboxes = page.locator("input[type=checkbox]").all()
        print(f"Checkboxes in Add User form: {len(checkboxes)}")
        for i, cb in enumerate(checkboxes):
            is_checked = cb.is_checked()
            # Get label text
            try:
                label = page.locator(f".el-checkbox__label").nth(i).text_content()
            except:
                label = "?"
            print(f"  checkbox[{i}]: checked={is_checked}, label={label!r}")

        # More detailed check via evaluate
        result = page.evaluate("""() => {
            const cbs = document.querySelectorAll('.el-dialog input[type=checkbox]');
            return Array.from(cbs).map(cb => {
                const label = cb.closest('.el-checkbox') ?
                    cb.closest('.el-checkbox').querySelector('.el-checkbox__label')?.textContent : '?';
                return {checked: cb.checked, label: label, id: cb.id, name: cb.name};
            });
        }""")
        print("\nDetailed checkbox state:")
        for item in result:
            print(f"  {item}")

        # Check form item labels
        form_items = page.locator(".el-dialog .el-form-item").all()
        print(f"\nForm items: {len(form_items)}")
        for fi in form_items:
            try:
                label = fi.locator(".el-form-item__label").text_content()
                print(f"  label: {label!r}")
            except:
                pass

        ctx.close()
        browser.close()
        print("Done.")


if __name__ == "__main__":
    run()
