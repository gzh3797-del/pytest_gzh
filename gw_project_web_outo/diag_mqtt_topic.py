"""Diagnostic: explore MQTT Topic and Parameter Selection page - find Base Topic field."""
from pages.login_page import LoginPage


def _nav_protocol(page, protocol: str, sub: str = None):
    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_mqtt_topic(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "MQTT", "Topic and Parameter Selection")
    print(f"\nURL: {page.url}")

    # All form items
    print("\n=== Form items ===")
    for fi in page.locator(".el-form-item").all():
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else "?"
            inputs = fi.locator("input").count()
            selects = fi.locator(".el-select").count()
            textareas = fi.locator("textarea").count()
            print(f"  label='{label}' inputs={inputs} selects={selects} textareas={textareas}")
        except Exception:
            pass

    # Specifically find Base Topic
    print("\n=== Base Topic form item ===")
    base_topic_fi = page.locator(".el-form-item").filter(has_text="Base Topic")
    print(f"  count: {base_topic_fi.count()}")
    if base_topic_fi.count() > 0:
        fi = base_topic_fi.first
        try:
            html = fi.evaluate("el => el.outerHTML")
            print(html[:1500])
        except Exception as e:
            print(f"  HTML error: {e}")

    # All inputs with attributes
    print("\n=== All inputs ===")
    for inp in page.locator("input").all():
        try:
            ph = inp.get_attribute("placeholder") or ""
            nm = inp.get_attribute("name") or ""
            typ = inp.get_attribute("type") or "text"
            val = inp.input_value() or ""
            cls = inp.get_attribute("class") or ""
            print(f"  placeholder='{ph}' name='{nm}' type='{typ}' value='{val[:40]}' class='{cls[:50]}'")
        except Exception:
            pass

    # All buttons
    print("\n=== Buttons ===")
    for btn in page.get_by_role("button").all():
        try:
            txt = btn.inner_text().strip()
            if txt:
                print(f"  '{txt}'")
        except Exception:
            pass

    assert True
