"""
Quick navigation exploration: login as admin, find all top-level nav names and URLs.
Then create a Device=view user and check what they can access.
"""
import sys
from playwright.sync_api import sync_playwright

BASE_URL = "https://192.168.2.199"
ADMIN_PW = "Admin@110001"

PERM_LABELS = [
    "User", "Device", "Data Log", "System Settings",
    "Protocol", "Alarm Log", "Maintenance", "Diagnostics", "Firmware Update",
]


def login(browser, username, password):
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    p.goto(BASE_URL + "/#/login")
    p.wait_for_load_state("networkidle")
    p.get_by_role("textbox", name="Enter User Name").fill(username)
    p.get_by_role("textbox", name="Enter Password").fill(password)
    p.get_by_role("button", name="Sign In").click()
    p.wait_for_load_state("networkidle")
    p.wait_for_timeout(1000)
    try:
        p.get_by_role("button", name="Accept").click(timeout=3000)
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(500)
    except Exception:
        pass
    try:
        p.get_by_role("button", name="Cancel").click(timeout=2000)
    except Exception:
        pass
    return ctx, p


def nav_user_mgmt(page, submenu):
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def create_role(page, role_name, perm_map):
    nav_user_mgmt(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    for lbl, val in perm_map.items():
        page.locator(".el-form-item").filter(has_text=lbl).locator(".el-select").click()
        page.wait_for_timeout(200)
        page.get_by_role("option", name=val, exact=True).click()
        page.wait_for_timeout(200)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def create_user(page, username, password, role):
    nav_user_mgmt(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def delete_role(page, role_name):
    nav_user_mgmt(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() > 0:
        row.get_by_role("button").last.click()
        page.get_by_role("button", name="Yes, continue").click()
        page.wait_for_timeout(500)


def delete_user(page, username):
    nav_user_mgmt(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() > 0:
        row.get_by_role("button").last.click()
        page.get_by_role("button", name="Yes, continue").click()
        page.wait_for_timeout(500)


with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)

    # ── Step 1: Admin - find all top-level navigation labels ──────────
    ctx_a, admin = login(browser, "admin", ADMIN_PW)
    admin.locator("header span").filter(has_text="AcuHMI").first.click()
    admin.wait_for_timeout(1500)
    # Find all clickable items in the dropdown menu
    menu_items = admin.locator(".el-menu > .el-menu-item, .el-menu > .el-sub-menu").all()
    print("=== Admin top-level nav items ===")
    for item in menu_items:
        try:
            txt = item.inner_text().strip().split('\n')[0]
            print(f"  [{txt}]")
        except:
            pass
    admin.screenshot(path="product_structure_testcase_regulation/explore/admin_nav_dropdown.png")
    admin.keyboard.press("Escape")
    admin.wait_for_timeout(300)

    # ── Step 2: Create Device=view user ──────────────────────────────
    perm_dv = {lbl: ("view" if lbl == "Device" else "none") for lbl in PERM_LABELS}
    create_role(admin, "xplDevV", perm_dv)
    create_user(admin, "xplDevV", ADMIN_PW, "xplDevV")

    perm_de = {lbl: ("edit" if lbl == "Device" else "none") for lbl in PERM_LABELS}
    create_role(admin, "xplDevE", perm_de)
    create_user(admin, "xplDevE", ADMIN_PW, "xplDevE")

    # ── Step 3: Login as Device=view user, explore navigation ────────
    ctx_dv, dv = login(browser, "xplDevV", ADMIN_PW)
    print(f"\nDevice=view URL after login: {dv.url}")

    # Try opening the nav dropdown
    try:
        dv.locator("header span").filter(has_text="AcuHMI").first.click()
        dv.wait_for_timeout(1500)
        dv.screenshot(path="product_structure_testcase_regulation/explore/dv_nav_dropdown.png")
        items = dv.locator(".el-menu > .el-menu-item, .el-menu > .el-sub-menu, [role=menuitem]").all()
        print("Device=view nav items:")
        for item in items:
            try:
                txt = item.inner_text().strip().split('\n')[0]
                if txt:
                    print(f"  [{txt}]")
            except:
                pass
        dv.keyboard.press("Escape")
    except Exception as e:
        print(f"Nav dropdown error: {e}")

    # Try clicking common nav labels
    for nav_label in ["Device", "Devices", "Device Management"]:
        try:
            dv.locator("header span").filter(has_text="AcuHMI").first.click()
            dv.wait_for_timeout(500)
            dv.get_by_text(nav_label, exact=True).first.click()
            dv.wait_for_load_state("networkidle")
            dv.wait_for_timeout(1000)
            print(f"Clicked [{nav_label}] → URL: {dv.url}")
            dv.screenshot(path=f"product_structure_testcase_regulation/explore/dv_{nav_label.replace(' ','_')}.png")
            # Check submenus
            subs = dv.locator("[role=menuitem]").all()
            for s in subs:
                try:
                    t = s.inner_text().strip()
                    if t:
                        print(f"  sub: [{t}]")
                except:
                    pass
        except Exception as e:
            print(f"  [{nav_label}] → {e}")

    ctx_dv.close()

    # ── Step 4: Login as Device=edit user ────────────────────────────
    ctx_de, de = login(browser, "xplDevE", ADMIN_PW)
    print(f"\nDevice=edit URL after login: {de.url}")
    for nav_label in ["Device", "Devices", "Device Management"]:
        try:
            de.locator("header span").filter(has_text="AcuHMI").first.click()
            de.wait_for_timeout(500)
            de.get_by_text(nav_label, exact=True).first.click()
            de.wait_for_load_state("networkidle")
            de.wait_for_timeout(1000)
            print(f"Device=edit Clicked [{nav_label}] → URL: {de.url}")
            de.screenshot(path=f"product_structure_testcase_regulation/explore/de_{nav_label.replace(' ','_')}.png")
        except Exception as e:
            print(f"  [{nav_label}] → {e}")
    ctx_de.close()

    # ── Cleanup ───────────────────────────────────────────────────────
    delete_user(admin, "xplDevV")
    delete_user(admin, "xplDevE")
    delete_role(admin, "xplDevV")
    delete_role(admin, "xplDevE")
    ctx_a.close()
    browser.close()

print("\n=== Done ===")
