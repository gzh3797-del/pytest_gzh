"""Diagnostic: explore Protocol field HTML and switching mechanism."""
from pages.login_page import LoginPage


def _nav_to_physical_devices(page):
    if "/#/physicalDevice" not in page.url or "addDevice" in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_diag_add_device2(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_physical_devices(page)
    page.get_by_role("button", name="Add Device").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Print Protocol form item's outerHTML
    print("\n=== Protocol form item HTML ===")
    proto_fi = page.locator(".el-form-item").filter(has_text="Protocol").first
    try:
        html = proto_fi.evaluate("el => el.outerHTML")
        print(html[:2000])
    except Exception as e:
        print(f"Error: {e}")

    # Print ALL elements inside Protocol form item
    print("\n=== All child elements in Protocol form item ===")
    for child in proto_fi.locator("*").all():
        try:
            tag = child.evaluate("el => el.tagName")
            cls = child.get_attribute("class") or ""
            txt = child.inner_text().strip()[:50]
            typ = child.get_attribute("type") or ""
            val = child.get_attribute("value") or ""
            if tag and cls:
                print(f"  <{tag}> class='{cls[:60]}' type='{typ}' value='{val}' text='{txt}'")
        except Exception:
            pass

    # Check for radio buttons
    print("\n=== Radio buttons on page ===")
    for rb in page.locator("input[type='radio']").all():
        try:
            val = rb.get_attribute("value") or ""
            checked = rb.is_checked()
            parent_txt = rb.locator("..").inner_text().strip()[:80]
            print(f"  radio value='{val}' checked={checked} parent='{parent_txt}'")
        except Exception:
            pass

    # Check el-radio groups
    print("\n=== El-radio-group / el-radio on page ===")
    for rb in page.locator(".el-radio, .el-radio-group, .el-radio-button").all():
        try:
            txt = rb.inner_text().strip()[:80]
            cls = rb.get_attribute("class") or ""
            print(f"  class='{cls[:60]}' text='{txt}'")
        except Exception:
            pass

    # Check tabs
    print("\n=== Tabs ===")
    for tab in page.locator(".el-tabs__item, .el-tab-pane").all():
        try:
            txt = tab.inner_text().strip()
            if txt:
                print(f"  tab: '{txt}'")
        except Exception:
            pass

    assert True
