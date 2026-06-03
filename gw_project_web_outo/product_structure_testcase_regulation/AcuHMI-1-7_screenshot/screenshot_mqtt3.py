# -*- coding: utf-8 -*-
"""
Capture MQTT sub-pages using hover to reveal the dropdown.
Element UI horizontal sub-menus show on hover, not click.
"""
import asyncio, re, sys
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
    for s in ['div.nav-item:has-text("AcuHMI-1-7")',
              'div.nav-item-menu:has-text("AcuHMI-1-7")']:
        try:
            els = await page.query_selector_all(s)
            for el in els:
                if "AcuHMI" in (await el.inner_text()):
                    await el.click(); await settle(page, 10000)
                    await page.wait_for_timeout(1200); break
        except Exception: pass
    try:
        els = await page.query_selector_all("li.left-nav-item")
        for el in els:
            if (await el.inner_text()).strip() == "Protocols":
                await el.click(); await settle(page, 5000); return True
    except Exception: pass
    return False

async def main():
    log("="*60)
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
        await open_protocols(page)

        log(f"  Protocols URL: {page.url}")

        # --- Method 1: Hover over MQTT tab ---
        log("\n  Method 1: Hover over MQTT")
        mqtt_el = await page.evaluate("""
            (function() {
                const subs = document.querySelectorAll('li.el-sub-menu');
                for (const sm of subs) {
                    const title = sm.querySelector('.el-sub-menu__title');
                    if (title && title.textContent.trim() === 'MQTT') {
                        const r = sm.getBoundingClientRect();
                        return {x: r.left + r.width/2, y: r.top + r.height/2};
                    }
                }
                return null;
            })()
        """)
        log(f"  MQTT element coords: {mqtt_el}")

        if mqtt_el:
            await page.mouse.move(mqtt_el['x'], mqtt_el['y'])
            await page.wait_for_timeout(800)

            # Screenshot with dropdown visible (don't scroll, don't escape)
            fname = f"{nxt()}_mqtt_hover.png"
            fpath = folder / fname
            await page.screenshot(path=str(fpath), full_page=False)
            _log.append({"path": str(fpath), "desc": "Protocols > MQTT (hover dropdown)"})
            log(f"  [SHOT] {fname}  (MQTT hover)")

            # Read sub-items while hovered
            subs_hover = await page.evaluate("""
                (function() {
                    const result = [];
                    // Look for popper/popup dropdown items
                    document.querySelectorAll(
                        'li.el-menu-item, .el-popper li, .el-sub-menu__popup li, ' +
                        'ul.el-menu--popup li'
                    ).forEach(el => {
                        if (el.offsetParent !== null) {
                            const r = el.getBoundingClientRect();
                            const txt = el.textContent.trim();
                            if (txt.length < 40 && txt.length > 1 &&
                                r.top > 50 && r.left > 150) {
                                result.push({text: txt, top: r.top, left: r.left,
                                             cls: el.className.slice(0,40)});
                            }
                        }
                    });
                    return result;
                })()
            """)
            log(f"  Sub-items while hovered: {subs_hover}")

            # Click each sub-item while keeping hover state
            seen = set()
            for sub in subs_hover:
                t = sub['text']
                if t in seen or len(t) > 40: continue
                # Skip items that are clearly not MQTT sub-items
                if t in ['Modbus', 'SNMP', 'BACnet/IP', 'AWS IoT', 'Azure IoT',
                         'Modbus Config', 'MQTT']: continue
                seen.add(t)
                log(f"    Clicking: '{t}'")
                clicked = await page.evaluate(f"""
                    (function() {{
                        const target = {repr(t)};
                        const els = document.querySelectorAll(
                            'li.el-menu-item, .el-popper li, ul.el-menu--popup li'
                        );
                        for (const el of els) {{
                            if (el.offsetParent !== null &&
                                el.textContent.trim() === target) {{
                                el.click(); return true;
                            }}
                        }}
                        return false;
                    }})()
                """)
                if clicked:
                    await settle(page, 4000)
                    url = page.url
                    log(f"    URL: {url}")
                    await shot(page, folder, f"mqtt_{clean(t)}", f"Protocols > MQTT > {t}")
                    # Re-hover MQTT to access next item
                    await page.mouse.move(mqtt_el['x'], mqtt_el['y'])
                    await page.wait_for_timeout(600)

        # --- Method 2: Brute-force URL guessing ---
        log("\n  Method 2: URL exploration")
        base = "https://192.168.2.199/#/protocols"
        candidate_urls = [
            f"{base}/mqtt/mqttConfig",
            f"{base}/mqtt/mqtt",
            f"{base}/mqtt/config",
            f"{base}/mqttBroker",
            f"{base}/mqtt",
        ]
        for url in candidate_urls:
            await page.goto(url, wait_until="networkidle", timeout=8000)
            await settle(page, 5000)
            current = page.url
            if current != "https://192.168.2.199/#/dashboard" and "mqtt" in current.lower():
                log(f"  Valid MQTT URL: {current}")
                await shot(page, folder, f"mqtt_url_{clean(current)}", f"MQTT - {current}")
                break
            else:
                log(f"  {url} -> redirected to {current}")

        await browser.close()

    log(f"\n  Done! {_n} screenshots.")

if __name__ == "__main__":
    asyncio.run(main())

