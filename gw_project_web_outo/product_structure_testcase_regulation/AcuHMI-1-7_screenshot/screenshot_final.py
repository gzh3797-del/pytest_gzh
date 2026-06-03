# -*- coding: utf-8 -*-
"""
AcuHMI-1-7 Final Pass – missing items:
  1. Protocols > MQTT tab
  2. Physical / Virtual Devices > Add Device dialog
  3. User Management > Add User dialog
  4. Modbus dropdown sub-pages
  5. Check Alarm / Data Log for any sub-menus
"""
import asyncio, re, sys, datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL = "https://192.168.2.199"
USERNAME = "admin"
PASSWORD = "Admin@110001"
OUT_DIR  = Path(r"C:\knowledge_base\AcuHMI-1-7_screenshot")
W, H     = 1920, 1080

DANGER = re.compile(
    r"delete|remove|restore|reset.factory|factory.reset|clear.all|erase|format|"
    r"reboot|restart|删除|恢复出厂|清除|格式化|重置|重启",
    re.IGNORECASE,
)
_n, _log = 0, []

def clean(s):
    s = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", s)
    return re.sub(r"\s+", "_", s.strip())[:60] or "unnamed"

def nxt():
    global _n; _n += 1
    return f"{_n:04d}"

def log(msg): print(msg, flush=True)

async def scroll_top(page):
    await page.evaluate("window.scrollTo(0,0)")
    await page.wait_for_timeout(300)

async def shot(page, folder, name, desc=""):
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}.png"
    fpath = folder / fname
    await scroll_top(page)
    await page.wait_for_timeout(500)
    try:
        await page.screenshot(path=str(fpath), full_page=False)
        _log.append({"path": str(fpath), "desc": desc or name})
        log(f"  [SHOT] {fname}  ({desc or name})")
        return fpath
    except Exception as e:
        log(f"  [WARN] {name}: {e}"); return None

async def settle(page, ms=8000):
    try: await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout: pass
    await page.wait_for_timeout(700)

async def login(page):
    await page.goto(BASE_URL, wait_until="networkidle", timeout=30000)
    for s in ['input[name="username"]', 'input[type="text"]']:
        try:
            await page.wait_for_selector(s, timeout=2000, state="visible")
            await page.fill(s, USERNAME); break
        except Exception: pass
    await page.fill('input[type="password"]', PASSWORD)
    for s in ['button[type="submit"]', 'button:has-text("Login")', 'button']:
        try:
            b = await page.query_selector(s)
            if b and await b.is_visible(): await b.click(); break
        except Exception: pass
    await settle(page, 15000)
    log("  [OK] Logged in")

async def close_dialog(page):
    for cs in [".ant-modal-close", ".el-dialog__headerbtn",
               'button:has-text("Cancel")', 'button:has-text("取消")',
               'button:has-text("Close")', "[aria-label='Close']"]:
        try:
            c = await page.query_selector(cs)
            if c and await c.is_visible():
                await c.click(timeout=1500)
                await page.wait_for_timeout(400)
                return
        except Exception: pass
    try: await page.keyboard.press("Escape")
    except Exception: pass
    await page.wait_for_timeout(400)

async def open_sys_settings(page):
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)
    for s in ['div.nav-item:has-text("AcuHMI-1-7")',
              'div.nav-item-menu:has-text("AcuHMI-1-7")']:
        try:
            els = await page.query_selector_all(s)
            for el in els:
                if "AcuHMI" in (await el.inner_text()):
                    await el.click()
                    await settle(page, 10000)
                    await page.wait_for_timeout(1200)
                    return True
        except Exception: pass
    return False

async def click_left(page, label):
    try:
        els = await page.query_selector_all("li.left-nav-item")
        for el in els:
            if (await el.inner_text()).strip() == label:
                await el.click(); await settle(page, 5000); return True
    except Exception: pass
    return False

# ── 1. MQTT tab under Protocols ───────────────────────────────────────────────
async def do_mqtt(page):
    log("\n=== 1. Protocols > MQTT ===")
    folder = OUT_DIR / "SysSettings_03_Protocols"

    if not await open_sys_settings(page): return
    if not await click_left(page, "Protocols"): return

    # MQTT is visible in the Protocols tab bar
    # Use JavaScript to find and click it
    result = await page.evaluate("""
        (function() {
            const els = document.querySelectorAll('*');
            for (const el of els) {
                if (el.textContent.trim() === 'MQTT' &&
                    el.offsetParent !== null &&
                    el.children.length === 0) {
                    el.click();
                    return true;
                }
            }
            return false;
        })()
    """)
    if result:
        await settle(page, 5000)
        await shot(page, folder, "tab_MQTT", "Protocols > MQTT")
        log("  [OK] MQTT captured")
    else:
        log("  [MISS] MQTT element not found")

    # Also capture Modbus sub-items by clicking Modbus dropdown
    log("\n  Modbus dropdown sub-pages...")
    result2 = await page.evaluate("""
        (function() {
            const els = document.querySelectorAll('*');
            for (const el of els) {
                const txt = el.textContent.trim();
                if ((txt === 'Modbus' || txt.startsWith('Modbus')) &&
                    el.offsetParent !== null &&
                    el.children.length <= 1) {
                    el.click();
                    return txt;
                }
            }
            return null;
        })()
    """)
    if result2:
        await page.wait_for_timeout(800)
        # Capture any dropdown items that appeared
        dropdown_items = await page.evaluate("""
            (function() {
                const result = [];
                document.querySelectorAll('li, a, span, div').forEach(el => {
                    if (el.offsetParent !== null) {
                        const txt = el.textContent.trim();
                        const cls = el.className || '';
                        if (txt.length < 40 && txt.length > 2 &&
                            !txt.includes('\\n') &&
                            (cls.includes('menu-item') || cls.includes('dropdown') ||
                             cls.includes('submenu') || cls.includes('item'))) {
                            const r = el.getBoundingClientRect();
                            if (r.top > 50 && r.top < 300 && r.left > 130) {
                                result.push(txt);
                            }
                        }
                    }
                });
                return [...new Set(result)];
            })()
        """)
        log(f"  Modbus dropdown items: {dropdown_items}")
        for item in dropdown_items:
            if DANGER.search(item) or item in ["Modbus"]: continue
            found = await page.evaluate(f"""
                (function() {{
                    const target = {repr(item)};
                    const els = document.querySelectorAll('*');
                    for (const el of els) {{
                        if (el.textContent.trim() === target && el.offsetParent !== null &&
                            el.children.length <= 1) {{
                            el.click(); return true;
                        }}
                    }}
                    return false;
                }})()
            """)
            if found:
                await settle(page, 5000)
                await shot(page, folder, f"tab_Modbus_{clean(item)}", f"Protocols > Modbus > {item}")

# ── 2. Physical / Virtual Devices – Add Device dialog ────────────────────────
async def do_add_device_dialogs(page):
    log("\n=== 2. Physical / Virtual Devices - Add Device Dialog ===")

    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    for label, folder_name in [
        ("Physical Devices", "02_Physical_Devices"),
        ("Virtual Devices",  "03_Virtual_Devices"),
    ]:
        folder = OUT_DIR / folder_name
        log(f"\n  {label}...")

        if not await click_left(page, label):
            log(f"  [ERR] Cannot nav to {label}"); continue

        await settle(page)

        # Screenshot current page state first
        await shot(page, folder, "before_add_dialog", f"{label} before Add")

        # Click Add Device / Add Virtual Device button
        btns = await page.query_selector_all("button")
        add_btn = None
        for btn in btns:
            if not await btn.is_visible(): continue
            txt = (await btn.inner_text()).strip()
            if "Add" in txt and not DANGER.search(txt):
                log(f"  Found button: '{txt}'")
                add_btn = btn; break

        if not add_btn:
            log(f"  [WARN] No Add button found for {label}"); continue

        await add_btn.click()
        await page.wait_for_timeout(1500)

        # Take a screenshot regardless — captures dialog if it appeared
        await shot(page, folder, "after_add_click", f"{label} - Add Device Dialog (after click)")

        # Also check specific dialog selectors
        dialog_found = False
        for ds in ["[role='dialog']", ".ant-modal", ".el-dialog", ".el-drawer",
                   ".modal", ".dialog", "[class*='dialog']", "[class*='modal']"]:
            try:
                d = await page.query_selector(ds)
                if d and await d.is_visible():
                    await shot(page, folder, f"dialog_{clean(ds)}", f"{label} - Add Dialog")
                    dialog_found = True
                    break
            except Exception: pass

        if not dialog_found:
            log(f"  [INFO] No standard dialog detected — screenshot taken regardless")

        await close_dialog(page)
        await page.wait_for_timeout(500)

# ── 3. User Management – Add User dialog ─────────────────────────────────────
async def do_add_user_dialog(page):
    log("\n=== 3. User Management - Add User Dialog ===")
    folder = OUT_DIR / "SysSettings_06_User_Management"

    if not await open_sys_settings(page): return
    if not await click_left(page, "User Management"): return

    # Navigate to User Configuration tab first
    await page.evaluate("""
        (function() {
            document.querySelectorAll('*').forEach(el => {
                if (el.textContent.trim() === 'User Configuration' &&
                    el.offsetParent !== null) { el.click(); }
            });
        })()
    """)
    await settle(page, 4000)

    # Click Add User button
    btns = await page.query_selector_all("button")
    for btn in btns:
        if not await btn.is_visible(): continue
        txt = (await btn.inner_text()).strip()
        if "Add User" in txt or (txt == "Add" and not DANGER.search(txt)):
            log(f"  Found: '{txt}'")
            await btn.click()
            await page.wait_for_timeout(1200)
            await shot(page, folder, "dialog_Add_User", "User Management - Add User Dialog")
            await close_dialog(page)
            break

# ── 4. Alarm and Data Log – check for sub-pages ───────────────────────────────
async def do_alarm_datalog_check(page):
    log("\n=== 4. Alarm / Data Log sub-pages check ===")

    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    for label, folder_name in [("Alarm", "05_Alarm"), ("Data Log", "06_Data_Log")]:
        folder = OUT_DIR / folder_name
        log(f"\n  {label}...")
        if not await click_left(page, label): continue
        await settle(page)

        # Check for secondary navigation (sub-pages or tabs)
        sub_items = await page.evaluate("""
            (function() {
                const result = [];
                document.querySelectorAll('li, a, span').forEach(el => {
                    if (el.offsetParent !== null) {
                        const txt = el.textContent.trim();
                        const cls = el.className || '';
                        if (txt.length < 40 && txt.length > 2 &&
                            !txt.includes('\\n') &&
                            (cls.includes('menu-item') || cls.includes('tab') ||
                             cls.includes('item'))) {
                            const r = el.getBoundingClientRect();
                            if (r.top > 50 && r.top < 200 && r.left > 130) {
                                result.push({text: txt, cls: cls.slice(0, 40)});
                            }
                        }
                    }
                });
                return result;
            })()
        """)
        log(f"  Sub-items: {[(x['text'], x['cls']) for x in sub_items[:8]]}")

        # Click any sub-items found that aren't breadcrumbs
        seen = set()
        skip_texts = {label, "AcuHMI-1-7", "Dashboard", "Physical Devices",
                      "Virtual Devices", "Web Devices", "Alarm", "Data Log"}
        for item in sub_items:
            t = item["text"]
            if t in seen or t in skip_texts or DANGER.search(t): continue
            seen.add(t)
            found = await page.evaluate(f"""
                (function() {{
                    const target = {repr(t)};
                    const els = document.querySelectorAll('*');
                    for (const el of els) {{
                        if (el.textContent.trim() === target && el.offsetParent !== null &&
                            el.children.length <= 1) {{
                            el.click(); return true;
                        }}
                    }}
                    return false;
                }})()
            """)
            if found:
                await settle(page, 4000)
                await shot(page, folder, f"sub_{clean(t)}", f"{label} > {t}")

def write_report():
    rp = OUT_DIR / "final_index.md"
    lines = [
        "# AcuHMI-1-7 Final Pass Screenshots",
        f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}",
        f"Screenshots: {_n}", "",
        "| # | File | Description |",
        "|---|------|-------------|",
    ]
    for i, e in enumerate(_log, 1):
        try: rel = Path(e["path"]).relative_to(OUT_DIR)
        except Exception: rel = e.get("path", "")
        lines.append(f"| {i} | `{rel}` | {e['desc']} |")
    rp.write_text("\n".join(lines), encoding="utf-8")
    log(f"  [OK] Report: {rp}")

async def main():
    log("=" * 60)
    log("  AcuHMI-1-7 Final Pass")
    log("=" * 60)
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", f"--window-size={W},{H}"],
        )
        ctx = await browser.new_context(
            ignore_https_errors=True, viewport={"width": W, "height": H}
        )
        page = await ctx.new_page()
        await login(page)

        await do_mqtt(page)
        await do_add_device_dialogs(page)
        await do_add_user_dialog(page)
        await do_alarm_datalog_check(page)

        await browser.close()
    write_report()
    log(f"\n  Final pass done! {_n} screenshots.")

if __name__ == "__main__":
    asyncio.run(main())

