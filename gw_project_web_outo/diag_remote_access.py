"""Diagnostic: explore Remote Access page structure."""
from pages.login_page import LoginPage


def test_diag_remote_access(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Navigate to System Settings → Remote Access
    page.locator("header span").filter(has_text="AcuHMI").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    page.locator(".el-menu-item").filter(has_text="Remote Access").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    print(f"\nURL: {page.url}")

    # All form items
    print("\n=== Form items ===")
    for fi in page.locator(".el-form-item").all():
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
            txt = fi.inner_text().strip().replace('\n', ' | ')[:100]
            print(f"  label='{label}' content='{txt}'")
        except Exception:
            pass

    # All buttons
    print("\n=== Buttons ===")
    for btn in page.get_by_role("button").all():
        try:
            txt = btn.inner_text().strip()
            cls = btn.get_attribute("class") or ""
            if txt:
                print(f"  '{txt}' class='{cls[:60]}'")
        except Exception:
            pass

    # Page text
    print("\n=== Page visible text ===")
    print(page.locator("main, .el-main, .content-wrapper").first.inner_text()[:800])

    assert True
