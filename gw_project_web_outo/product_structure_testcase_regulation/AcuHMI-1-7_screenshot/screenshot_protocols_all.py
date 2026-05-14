# -*- coding: utf-8 -*-
"""Capture all hover sub-pages for Modbus, BACnet/IP, AWS IoT, Azure IoT."""
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

async def hover_and_capture(page, folder, protocol_name, skip_texts):
    """Hover over a protocol tab and capture its sub-items."""
    log(f"\n  === {protocol_name} ===")

    # Get element coords
    coords = await page.evaluate(f"""
        (function() {{
            const target = {repr(protocol_name)};
            const subs = document.querySelectorAll('li.el-sub-menu, li.el-menu-item');
            for (const el of subs) {{
                const title = el.querySelector('.el-sub-menu__title') || el;
                if (title.textContent.trim() === target && el.offsetParent !== null) {{
                    const r = el.getBoundingClientRect();
                    return {{x: r.left + r.width/2, y: r.top + r.height/2,
                             isSub: el.classList.contains('el-sub-menu')}};
                }}
            }}
            return null;
        }})()
    """)
    if not coords:
        log(f"  [MISS] {protocol_name} not found")
        return []

    log(f"  Coords: {coords}")

    if coords.get('isSub'):
        # Hover to reveal sub-items
        await page.mouse.move(coords['x'], coords['y'])
        await page.wait_for_timeout(700)

        subs = await page.evaluate(f"""
            (function() {{
                const result = [];
                document.querySelectorAll('li.el-menu-item').forEach(el => {{
                    if (el.offsetParent !== null) {{
                        const r = el.getBoundingClientRect();
                        const txt = el.textContent.trim();
                        // Items appearing below the tab bar (dropdown)
                        if (r.top > 150 && r.left > 150 && txt.length < 50) {{
                            result.push({{text: txt, top: r.top, left: r.left}});
                        }}
                    }}
                }});
                return result;
            }})()
        """)
        log(f"  Sub-items: {[s['text'] for s in subs]}")
        captured = []
        seen = set()
        for sub in subs:
            t = sub['text']
            if t in seen or t in skip_texts: continue
            seen.add(t)
            log(f"    Clicking: '{t}'")
            clicked = await page.evaluate(f"""
                (function() {{
                    const target = {repr(t)};
                    const els = document.querySelectorAll('li.el-menu-item');
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
                await shot(page, folder, f"{clean(protocol_name)}_{clean(t)}",
                           f"Protocols > {protocol_name} > {t}")
                captured.append(t)
                # Re-hover
                await page.mouse.move(coords['x'], coords['y'])
                await page.wait_for_timeout(600)
        return captured
    else:
        # Direct click (el-menu-item)
        await page.mouse.click(coords['x'], coords['y'])
        await settle(page, 4000)
        await shot(page, folder, clean(protocol_name), f"Protocols > {protocol_name}")
        return [protocol_name]

async def main():
    log("="*60 + "\n  Protocol Sub-pages (Modbus, BACnet, AWS, Azure)\n" + "="*60)
    folder = OUT_DIR / "SysSettings_03_Protocols"

    # Known skip texts (already captured or navigation items)
    SKIP = {'AcuHMI-1-7', 'Protocols', 'Modbus', 'SNMP', 'BACnet/IP',
            'MQTT', 'AWS IoT', 'Azure IoT', 'Modbus Config',
            'General', 'User Credential', 'SSL/TLS',
            'Last Will and Testament', 'Topic and Parameter Selection'}

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
        log(f"  URL: {page.url}")

        # Check each protocol tab
        for proto in ["Modbus", "BACnet/IP", "AWS IoT", "Azure IoT"]:
            await open_protocols(page)   # navigate back each time
            result = await hover_and_capture(page, folder, proto, SKIP)
            log(f"  {proto} -> {result}")

    log(f"\n  Done! {_n} screenshots.")

if __name__ == "__main__":
    asyncio.run(main())

