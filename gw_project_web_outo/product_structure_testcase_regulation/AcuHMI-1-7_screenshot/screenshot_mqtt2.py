# -*- coding: utf-8 -*-
"""Capture MQTT sub-pages: expand MQTT el-sub-menu, click each sub-item."""
import asyncio, re, sys, datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL = "https://192.168.2.199"
USERNAME = "admin"; PASSWORD = "Admin@110001"
OUT_DIR  = Path(r"C:\knowledge_base\AcuHMI-1-7_screenshot")
W, H = 1920, 1080
_n, _log = 0, []

def clean(s):
    s = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", s)
    return re.sub(r"\s+", "_", s.strip())[:60] or "unnamed"
def nxt():
    global _n; _n += 1; return f"{_n:04d}"
def log(msg): print(msg, flush=True)

async def settle(page, ms=6000):
    try: await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout: pass
    await page.wait_for_timeout(600)

async def shot(page, folder, name, desc=""):
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}.png"
    fpath = folder / fname
    await page.evaluate("window.scrollTo(0,0)")
    await page.wait_for_timeout(400)
    await page.screenshot(path=str(fpath), full_page=False)
    _log.append({"path": str(fpath), "desc": desc or name})
    log(f"  [SHOT] {fname}  ({desc or name})")
    return fpath

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
    await settle(page, 15000); log("  [OK] Logged in")

async def open_protocols(page):
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)
    # Open System Settings
    for s in ['div.nav-item:has-text("AcuHMI-1-7")',
              'div.nav-item-menu:has-text("AcuHMI-1-7")']:
        try:
            els = await page.query_selector_all(s)
            for el in els:
                if "AcuHMI" in (await el.inner_text()):
                    await el.click(); await settle(page, 10000)
                    await page.wait_for_timeout(1200); break
        except Exception: pass
    # Click Protocols
    try:
        els = await page.query_selector_all("li.left-nav-item")
        for el in els:
            if (await el.inner_text()).strip() == "Protocols":
                await el.click(); await settle(page, 5000); return True
    except Exception: pass
    return False

async def main():
    log("="*60 + "\n  MQTT Sub-pages Capture\n" + "="*60)
    folder = OUT_DIR / "SysSettings_03_Protocols"

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", f"--window-size={W},{H}"]
        )
        ctx = await browser.new_context(ignore_https_errors=True,
                                         viewport={"width": W, "height": H})
        page = await ctx.new_page()
        await login(page)

        if not await open_protocols(page):
            log("[ERR] Cannot open Protocols"); await browser.close(); return

        # Log the current URL to understand URL pattern
        log(f"  Protocols URL: {page.url}")

        # Screenshot current state (SNMP should be first active)
        await shot(page, folder, "protocols_landing", "Protocols - Landing page")

        # Step 1: Find MQTT el-sub-menu title and click to expand
        mqtt_title = await page.evaluate("""
            (function() {
                // Find the el-sub-menu that contains 'MQTT' text
                const submenus = document.querySelectorAll('li.el-sub-menu');
                for (const sm of submenus) {
                    const title = sm.querySelector('.el-sub-menu__title');
                    if (title && title.textContent.trim().includes('MQTT')) {
                        title.click();
                        return 'clicked: ' + title.textContent.trim();
                    }
                }
                return 'not found';
            })()
        """)
        log(f"  MQTT title click: {mqtt_title}")
        await page.wait_for_timeout(800)

        # Step 2: Capture expanded state + list sub-items
        await shot(page, folder, "mqtt_dropdown_open", "Protocols > MQTT (dropdown open)")

        # Step 3: Find sub-items in the expanded MQTT sub-menu
        mqtt_subs = await page.evaluate("""
            (function() {
                const result = [];
                const submenus = document.querySelectorAll('li.el-sub-menu');
                for (const sm of submenus) {
                    const title = sm.querySelector('.el-sub-menu__title');
                    if (title && title.textContent.trim().includes('MQTT')) {
                        // Look for sub-menu items
                        const items = sm.querySelectorAll('li.el-menu-item, li.el-sub-menu');
                        for (const item of items) {
                            if (item.offsetParent !== null) {
                                result.push(item.textContent.trim());
                            }
                        }
                        // Also check popper/dropdown if it appeared outside the LI
                        break;
                    }
                }
                // Also check for any newly visible items at the right position
                document.querySelectorAll('li.el-menu-item').forEach(el => {
                    if (el.offsetParent !== null) {
                        const r = el.getBoundingClientRect();
                        // Items that appeared in the dropdown area (right side, below tab bar)
                        if (r.top > 130 && r.left > 400) {
                            result.push('dropdown: ' + el.textContent.trim());
                        }
                    }
                });
                return [...new Set(result)];
            })()
        """)
        log(f"  MQTT sub-items: {mqtt_subs}")

        # Step 4: Try to click each MQTT sub-item
        for sub in mqtt_subs:
            if not sub or "dropdown: " not in sub:
                continue
            txt = sub.replace("dropdown: ", "").strip()
            if not txt or len(txt) > 40: continue
            log(f"  Clicking MQTT sub: '{txt}'")
            found = await page.evaluate(f"""
                (function() {{
                    const target = {repr(txt)};
                    const items = document.querySelectorAll('li.el-menu-item');
                    for (const el of items) {{
                        if (el.offsetParent !== null &&
                            el.textContent.trim() === target) {{
                            el.click(); return true;
                        }}
                    }}
                    return false;
                }})()
            """)
            if found:
                await settle(page, 4000)
                url = page.url
                log(f"    URL: {url}")
                await shot(page, folder, f"mqtt_sub_{clean(txt)}", f"Protocols > MQTT > {txt}")
                # Re-open MQTT dropdown for next item
                await page.evaluate("""
                    (function() {
                        const submenus = document.querySelectorAll('li.el-sub-menu');
                        for (const sm of submenus) {
                            const title = sm.querySelector('.el-sub-menu__title');
                            if (title && title.textContent.trim().includes('MQTT')) {
                                title.click(); return;
                            }
                        }
                    })()
                """)
                await page.wait_for_timeout(600)

        # Step 5: Check URLs to understand the pattern, try direct URL navigation
        # Navigate to Modbus Config to get URL format
        log("\n  Checking URL patterns...")
        await page.evaluate("""
            (function() {
                const submenus = document.querySelectorAll('li.el-sub-menu');
                for (const sm of submenus) {
                    const title = sm.querySelector('.el-sub-menu__title');
                    if (title && title.textContent.trim().includes('Modbus')) {
                        title.click(); return;
                    }
                }
            })()
        """)
        await page.wait_for_timeout(600)
        # Click Modbus Config
        await page.evaluate("""
            (function() {
                document.querySelectorAll('li.el-menu-item').forEach(el => {
                    if (el.offsetParent !== null &&
                        el.textContent.trim() === 'Modbus Config') {
                        el.click();
                    }
                });
            })()
        """)
        await settle(page, 4000)
        modbus_url = page.url
        log(f"  Modbus Config URL: {modbus_url}")

        # Derive MQTT URL from Modbus URL
        if "modbus" in modbus_url.lower():
            mqtt_url = re.sub(r'modbus[^/]*', 'mqtt', modbus_url, flags=re.IGNORECASE)
            if mqtt_url != modbus_url:
                log(f"  Trying MQTT URL: {mqtt_url}")
                await page.goto(mqtt_url, wait_until="networkidle", timeout=10000)
                await settle(page)
                current = page.url
                log(f"  Result URL: {current}")
                if "mqtt" in current.lower() or current != modbus_url:
                    await shot(page, folder, "mqtt_via_url", f"Protocols > MQTT (direct URL)")
                    # Now get the URL and check for sub-pages
                    all_menu = await page.evaluate("""
                        (function() {
                            const r = [];
                            document.querySelectorAll('li.el-menu-item, li.el-sub-menu').forEach(el => {
                                if (el.offsetParent !== null) {
                                    r.push({
                                        text: el.textContent.trim().slice(0,30),
                                        cls: el.className.slice(0,40),
                                        top: el.getBoundingClientRect().top
                                    });
                                }
                            });
                            return r;
                        })()
                    """)
                    log(f"  Menu items on MQTT page: {all_menu}")

        await browser.close()

    log(f"\n  Done! {_n} screenshots.")

if __name__ == "__main__":
    asyncio.run(main())

