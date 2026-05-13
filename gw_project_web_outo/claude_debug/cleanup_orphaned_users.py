"""
Delete orphaned test users left over from failed test runs.
Run this before re-running UC tests.
"""
import sys
sys.path.insert(0, r"C:\autotest_local\autotest\gw_project_web_outo")

from playwright.sync_api import sync_playwright
from config.settings import BASE_URL

_ADMIN = "admin"
_PWD   = "Admin@110001"

ORPHANED = [
    "uc203_01",
    "uc04lock1",
    "uc04lock2",
    "uc041_1",
    "uc041_2",
    "uc041_3",
    "uc041_4",
    "uc041_5",
]

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True, slow_mo=0)
    ctx = browser.new_context(ignore_https_errors=True)
    page = ctx.new_page()

    page.goto(BASE_URL + "/#/login")
    page.wait_for_load_state("networkidle")
    page.get_by_role("textbox", name="Enter User Name").fill(_ADMIN)
    page.get_by_role("textbox", name="Enter Password").fill(_PWD)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    for btn in ["Accept", "Cancel"]:
        try:
            page.get_by_role("button", name=btn).click(timeout=2000)
            page.wait_for_timeout(500)
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

    for username in ORPHANED:
        row = page.locator("tbody").get_by_role("row").filter(has_text=username)
        if row.count() == 0:
            print(f"[skip] {username} not found")
            continue
        row.get_by_role("button").last.click()
        page.wait_for_timeout(500)
        try:
            page.get_by_role("button", name="Yes, continue").click(timeout=3000)
            page.wait_for_timeout(500)
            print(f"[deleted] {username}")
        except Exception as e:
            print(f"[error] {username}: {e}")
        # Re-navigate to refresh table
        page.get_by_role("menuitem", name="User Configuration").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    rows = page.locator("tbody").get_by_role("row").all()
    print(f"\n[done] Remaining users: {len(rows)}")
    for row in rows:
        try:
            cells = row.locator("td").all()
            print(f"  {[c.inner_text().strip()[:30] for c in cells[:3]]}")
        except Exception:
            pass

    ctx.close()
    browser.close()

print("[cleanup] Done.")
