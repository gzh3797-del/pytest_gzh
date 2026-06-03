# -*- coding: utf-8 -*-
"""
Final targeted pass:
  1. MQTT tab in Protocols (dump DOM to find exact element, then click)
  2. Full-page screenshots of Physical/Virtual Add Device forms
  3. Virtual Devices Add Virtual Device form full-page
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
_n, _log = 0, []

def clean(s):
    s = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", s)
    return re.sub(r"\s+", "_", s.strip())[:60] or "unnamed"

def nxt():
    global _n; _n += 1
    return f"{_n:04d}"

def log(msg): print(msg, flush=True)

async def settle(page, ms=8000):
    try: await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout: pass
    await page.wait_for_timeout(700)

async def shot(page, folder, name, desc="", fp=False):
    folder.mkdir(parents=True, exist_ok=True)
    suffix = "_fp" if fp else ""
    fname = f"{nxt()}_{clean(name)}{suffix}.png"
    fpath = folder / fname
    await page.evaluate("window.scrollTo(0,0)")
    await page.wait_for_timeout(400)
    try:
        await page.screenshot(path=str(fpath), full_page=fp)
        _log.append({"path": str(fpath), "desc": desc or name})
        log(f"  [SHOT] {fname}  ({desc or name})")
        return fpath
    except Exception as e:
        log(f"  [WARN] {e}"); return None

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

# ── 1. MQTT – dump DOM then click ────────────────────────────────────────────
async def do_mqtt(page):
    log("\n=== MQTT Protocol Tab ===")
    folder = OUT_DIR / "SysSettings_03_Protocols"

    if not await open_sys_settings(page): return
    if not await click_left(page, "Protocols"): return

    # Dump all elements containing "MQTT"
    mqtt_info = await page.evaluate("""
        (function() {
            const result = [];
            document.querySelectorAll('*').forEach(el => {
                const txt = el.textContent.trim();
                if (txt === 'MQTT' || txt.includes('MQTT')) {
                    const r = el.getBoundingClientRect();
                    result.push({
                        tag: el.tagName,
                        cls: (el.className || '').slice(0, 80),
                        text: txt.slice(0, 30),
                        childCount: el.children.length,
                        visible: el.offsetParent !== null,
                        top: r.top, left: r.left, w: r.width, h: r.height
                    });
                }
            });
            return result;
        })()
    """)
    log(f"  MQTT elements in DOM: {len(mqtt_info)}")
    for info in mqtt_info:
        log(f"    {info}")

    # Try clicking MQTT — use index-based approach since text matcher failed
    clicked = await page.evaluate("""
        (function() {
            // Find all elements whose text is exactly 'MQTT' or contains 'MQTT' leaf node
            const candidates = [];
            document.querySelectorAll('li, span, a, div').forEach(el => {
                if (el.offsetParent !== null) {
                    // Check if this element or its immediate span/text contains MQTT
                    const own = el.childNodes;
                    let hasTextMQTT = false;
                    for (const n of own) {
                        if (n.nodeType === 3 && n.textContent.trim() === 'MQTT') {
                            hasTextMQTT = true; break;
                        }
                    }
                    if (hasTextMQTT || el.textContent.trim() === 'MQTT') {
                        candidates.push(el);
                    }
                }
            });
            if (candidates.length > 0) {
                candidates[0].click();
                return candidates[0].className + '|' + candidates[0].tagName;
            }
            return null;
        })()
    """)
    log(f"  MQTT click result: {clicked}")

    if clicked:
        await settle(page, 5000)
        await shot(page, folder, "tab_MQTT", "Protocols > MQTT")
    else:
        # Navigate directly via URL hash if possible
        current_url = page.url
        log(f"  Current URL: {current_url}")
        # Try hash-based navigation for MQTT
        mqtt_url = current_url.replace("/modbus", "/mqtt").replace(
            "dateTime", "mqtt"
        )
        if mqtt_url != current_url:
            await page.goto(mqtt_url, wait_until="networkidle", timeout=10000)
            await settle(page)
            await shot(page, folder, "tab_MQTT", "Protocols > MQTT (via URL)")

        # As fallback: just dump all visible menu items
        all_menu = await page.evaluate("""
            (function() {
                const r = [];
                document.querySelectorAll('li.el-menu-item, li.el-submenu__title').forEach(el => {
                    if (el.offsetParent !== null) {
                        r.push({text: el.textContent.trim(), cls: el.className.slice(0,40)});
                    }
                });
                return r;
            })()
        """)
        log(f"  All menu items (el-menu-item): {all_menu}")

# ── 2. Physical Devices Add Device – full-page ────────────────────────────────
async def do_add_device_fullpage(page):
    log("\n=== Physical / Virtual Devices Add Device – full-page ===")

    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    for label, folder_name in [
        ("Physical Devices", "02_Physical_Devices"),
        ("Virtual Devices",  "03_Virtual_Devices"),
    ]:
        folder = OUT_DIR / folder_name
        log(f"\n  {label}...")

        if not await click_left(page, label): continue
        await settle(page)

        btns = await page.query_selector_all("button")
        for btn in btns:
            if not await btn.is_visible(): continue
            txt = (await btn.inner_text()).strip()
            if "Add" in txt:
                log(f"  Clicking: '{txt}'")
                await btn.click()
                await settle(page, 5000)
                # Full-page screenshot to capture entire form
                await shot(page, folder, "add_device_form", f"{label} - Add Device Form", fp=True)
                # Also viewport screenshot
                await shot(page, folder, "add_device_form_vp", f"{label} - Add Device Form (viewport)")

                # Go back
                try:
                    cancel = await page.query_selector('button:has-text("Cancel")')
                    if cancel and await cancel.is_visible():
                        await cancel.click()
                    else:
                        await page.go_back()
                except Exception:
                    await page.go_back()
                await settle(page, 3000)
                break

def write_report():
    rp = OUT_DIR / "mqtt_final_index.md"
    lines = [
        "# MQTT & Add Device Forms",
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
    log(f"  [OK]: {rp}")

async def main():
    log("=" * 60)
    log("  AcuHMI-1-7 MQTT + Full Page Forms")
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
        await do_add_device_fullpage(page)
        await browser.close()
    write_report()
    log(f"\n  Done! {_n} screenshots.")

if __name__ == "__main__":
    asyncio.run(main())

