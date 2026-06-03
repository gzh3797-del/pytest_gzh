"""Diagnostic: dump Date & Time page structure."""
from pages.login_page import LoginPage


def test_diag_datetime(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    print(f"\nURL: {page.url}")

    print("\n=== All form items ===")
    for i, fi in enumerate(page.locator(".el-form-item").all()):
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
            txt = fi.inner_text().strip().replace('\n', ' | ')[:120]
            print(f"  [{i}] label='{label}' content='{txt}'")
        except Exception:
            pass

    print("\n=== All buttons ===")
    for btn in page.get_by_role("button").all():
        try:
            txt = btn.inner_text().strip()
            if txt:
                print(f"  '{txt}'")
        except Exception:
            pass

    print("\n=== All inputs ===")
    for i, inp in enumerate(page.locator("input").all()):
        try:
            ph = inp.get_attribute("placeholder") or ""
            val = inp.input_value()
            typ = inp.get_attribute("type") or "text"
            disabled = inp.is_disabled()
            print(f"  [{i}] type={typ} placeholder='{ph}' value='{val}' disabled={disabled}")
        except Exception:
            pass

    assert True
