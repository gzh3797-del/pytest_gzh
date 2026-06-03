"""Reset password policy to default values for testing."""
import sys
from playwright.sync_api import sync_playwright

BASE_URL = "https://192.168.2.199"
ADMIN_PW = "Admin@110001"

DEFAULTS = {
    "Enter Password History":              "1",
    "Enter Minimum Password Age":          "0",
    "Enter Password Expires":              "0",
    "Enter Minimum Password Length":       "8",
    "Enter Grace Period":                  "0",
    "Enter Maximum Failed Attempts":       "0",
    "Enter Failed Login Attempt Window":   "0",
    "Enter Failed Login Wait":             "0",
}


def login(br):
    ctx = br.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    p.goto(BASE_URL + "/#/login")
    p.wait_for_load_state("networkidle")
    p.get_by_role("textbox", name="Enter User Name").fill("admin")
    p.get_by_role("textbox", name="Enter Password").fill(ADMIN_PW)
    p.get_by_role("button", name="Sign In").click()
    p.wait_for_load_state("networkidle")
    p.wait_for_timeout(2000)
    try:
        p.get_by_role("button", name="Accept").click(timeout=2000)
        p.wait_for_load_state("networkidle")
    except Exception:
        pass
    try:
        p.get_by_role("button", name="Cancel").click(timeout=2000)
    except Exception:
        pass
    return ctx, p


with sync_playwright() as pw:
    br = pw.chromium.launch(headless=True, slow_mo=300)
    ctx, page = login(br)

    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Password Policy").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    for placeholder, value in DEFAULTS.items():
        try:
            inp = page.get_by_placeholder(placeholder)
            inp.click()
            inp.fill(value)
            sys.stdout.write(f"  Set [{placeholder}] = {value}\n")
            sys.stdout.flush()
        except Exception as e:
            sys.stdout.write(f"  ERROR [{placeholder}]: {e}\n")
            sys.stdout.flush()

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(2000)

    texts = page.locator(".el-message, .el-notification__content").all_inner_texts()
    sys.stdout.write(f"Save result: {texts}\n")
    sys.stdout.flush()

    sys.stdout.write("\nVerifying:\n")
    for placeholder, expected in DEFAULTS.items():
        try:
            val = page.get_by_placeholder(placeholder).input_value()
            ok = "OK" if val == expected else f"FAIL got={val!r}"
            sys.stdout.write(f"  [{placeholder[:35]}]: {ok}\n")
        except Exception as e:
            sys.stdout.write(f"  [{placeholder[:35]}]: ERROR {e}\n")
        sys.stdout.flush()

    ctx.close()
    br.close()
    sys.stdout.write("Done!\n")
    sys.stdout.flush()
