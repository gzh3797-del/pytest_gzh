"""Diagnostic: Azure IoT - select device then test Connection String validation."""
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


def test_diag_azure(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "Azure IoT")

    # Enable Azure IoT
    enable_item = page.locator(".el-form-item").filter(has_text="Azure IoT Enable").first
    if "is-checked" not in (enable_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        enable_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Inspect device selection tables
    print("\n=== Empty-label form items (device tables) ===")
    for i, fi in enumerate(page.locator(".el-form-item").all()):
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else ""
            if label == "":
                cbs = fi.locator("input[type='checkbox']").count()
                rows = fi.locator("tbody tr").count()
                print(f"  form-item[{i}] label='' checkboxes={cbs} tbody_rows={rows}")
                if rows > 0:
                    for j, row in enumerate(fi.locator("tbody tr").all()[:3]):
                        txt = row.inner_text().strip().replace('\n', ' | ')
                        print(f"    row[{j}]: '{txt[:80]}'")
        except Exception:
            pass

    # Select first device in first device table
    print("\n=== Selecting first device ===")
    for fi in page.locator(".el-form-item").all():
        try:
            label = fi.locator("label").first.inner_text().strip() if fi.locator("label").count() > 0 else ""
            if label == "" and fi.locator("tbody tr").count() > 0:
                first_row = fi.locator("tbody tr").first
                cb = first_row.locator("input[type='checkbox']").first
                if not cb.is_checked():
                    first_row.locator(".el-checkbox").first.click()
                    page.wait_for_timeout(300)
                    print(f"  Checked first device: '{first_row.inner_text().strip()[:60]}'")
                break
        except Exception:
            pass

    # Test 3 invalid Connection Strings
    test_cases = [
        ("missing_HostName",        "DeviceId=myDevice;SharedAccessKey=abc123=="),
        ("missing_DeviceId",        "HostName=myhub.azure-devices.net;SharedAccessKey=abc123=="),
        ("missing_SharedAccessKey", "HostName=myhub.azure-devices.net;DeviceId=myDevice"),
        ("fully_invalid",           "this is not a valid connection string!!!"),
    ]
    conn_inp = page.locator(".el-form-item").filter(has_text="Primary Connection String").first.locator("input").first

    for case_name, conn_str in test_cases:
        conn_inp.fill(conn_str)
        page.wait_for_timeout(200)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(800)

        field_errors = page.locator(".el-form-item__error").count()
        msg_errors = page.locator(".el-message--error").count()
        error_texts = []
        for e in page.locator(".el-form-item__error").all():
            try:
                if e.is_visible():
                    error_texts.append(e.inner_text().strip())
            except Exception:
                pass
        for e in page.locator(".el-message--error").all():
            try:
                error_texts.append(e.inner_text().strip())
            except Exception:
                pass
        print(f"\n  [{case_name}]")
        print(f"    field_errors={field_errors} msg_errors={msg_errors}")
        print(f"    error texts: {error_texts}")

    assert True
