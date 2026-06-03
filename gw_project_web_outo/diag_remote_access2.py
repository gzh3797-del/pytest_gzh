"""Diagnostic: Remote Access full page content after registration."""
from pages.login_page import LoginPage


def _nav_to_remote_access(page):
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)
    ra = page.locator(".el-menu-item").filter(has_text="Remote Access").first
    ra.click(force=True)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def test_diag_remote_access2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_remote_access(page)

    # Enable Remote Access
    enable_item = page.locator(".el-form-item").filter(has_text="Remote Access Enable").first
    enable_radio = enable_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (enable_radio.get_attribute("class") or ""):
        enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        _nav_to_remote_access(page)

    # Manual Register
    reg_btn = page.get_by_role("button", name="Manual Register")
    if reg_btn.count() == 0:
        reg_btn = page.locator("button").filter(has_text="Manual Register")
    if reg_btn.count() > 0:
        reg_btn.first.click()
        page.wait_for_timeout(1000)
        for btn_name in ["Yes, continue", "Yes,continue", "Yes", "Confirm", "确认", "OK"]:
            btn = page.get_by_role("button", name=btn_name)
            if btn.count() > 0:
                try:
                    if btn.first.is_visible():
                        btn.first.click()
                        page.wait_for_timeout(800)
                        break
                except Exception:
                    pass
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)

    _nav_to_remote_access(page)
    page.wait_for_timeout(2000)

    # Click Refresh Status
    refresh_btn = page.locator("button").filter(has_text="Refresh Status")
    if refresh_btn.count() > 0:
        refresh_btn.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)

    print(f"\nURL: {page.url}")

    # Full page text
    print("\n=== Full page body text ===")
    try:
        txt = page.locator("main, .el-main, .content-wrapper, .app-main, #app").first.inner_text()
        print(txt[:2000])
    except Exception:
        print(page.locator("body").inner_text()[:2000])

    # All elements containing "status" or "url" or "online"
    print("\n=== Elements with status/url/online text ===")
    for kw in ["status", "Status", "online", "URL", "url", "http"]:
        els = page.locator(f"*:has-text('{kw}')").all()
        for el in els[:5]:
            try:
                tag = el.evaluate("el => el.tagName")
                cls = el.get_attribute("class") or ""
                txt = el.inner_text().strip().replace('\n', ' | ')[:100]
                if txt and tag not in ["HTML", "BODY", "DIV", "SECTION", "MAIN", "ARTICLE"]:
                    print(f"  tag={tag} class='{cls[:40]}' text='{txt}'")
            except Exception:
                pass

    # All td / span / p elements
    print("\n=== td/span/p elements with content ===")
    for sel in ["td", "span.value", "p", ".info-value", ".field-value"]:
        for el in page.locator(sel).all()[:20]:
            try:
                txt = el.inner_text().strip()
                if txt and len(txt) < 200:
                    print(f"  {sel}: '{txt}'")
            except Exception:
                pass

    assert True
