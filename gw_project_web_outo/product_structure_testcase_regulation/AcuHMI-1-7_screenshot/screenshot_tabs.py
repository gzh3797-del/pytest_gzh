# -*- coding: utf-8 -*-
"""
AcuHMI-1-7 Pass 3 – System Settings tabs + Add Device dialogs
Uses same navigation approach as the supplement script (which worked).
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

# System Settings > System Settings tabs (from screenshot 0001)
SYS_SETTINGS_TABS = [
    "Date & Time", "Network", "Access Control", "Email",
    "Alarm Notification", "Certificate Management",
    "Configuration Management", "Remote Access",
]
# Maintenance tabs (from screenshot 0007)
MAINTENANCE_TABS = ["System Status", "Event Log"]

DANGER = re.compile(
    r"delete|remove|restore|reset.factory|factory.reset|clear.all|erase|format|"
    r"reboot|restart|删除|恢复出厂|清除|格式化|重置|重启",
    re.IGNORECASE,
)

_n   = 0
_log = []

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
        log(f"  [WARN] shot({name}): {e}")
        return None

async def settle(page, ms=8000):
    try:
        await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout:
        pass
    await page.wait_for_timeout(800)

async def login(page):
    log("  Logging in...")
    await page.goto(BASE_URL, wait_until="networkidle", timeout=30000)
    for s in ['input[name="username"]', 'input[type="text"]']:
        try:
            await page.wait_for_selector(s, timeout=2000, state="visible")
            await page.fill(s, USERNAME); break
        except Exception: pass
    await page.fill('input[type="password"]', PASSWORD)
    for s in ['button[type="submit"]', 'button:has-text("Login")',
              'button:has-text("登录")', 'button']:
        try:
            b = await page.query_selector(s)
            if b and await b.is_visible(): await b.click(); break
        except Exception: pass
    await settle(page, 15000)
    log("  [OK] Logged in")

async def close_dialog(page):
    for cs in [".ant-modal-close", ".el-dialog__headerbtn",
               'button:has-text("Cancel")', 'button:has-text("取消")',
               'button:has-text("Close")', 'button:has-text("关闭")',
               "[aria-label='Close']", "[aria-label='close']"]:
        try:
            c = await page.query_selector(cs)
            if c and await c.is_visible():
                await c.click(timeout=1500)
                await page.wait_for_timeout(500)
                return
        except Exception: pass
    try: await page.keyboard.press("Escape")
    except Exception: pass
    await page.wait_for_timeout(400)

# ── Exact same navigation as supplement.py (which confirmed working) ─────────
async def open_system_settings(page):
    """Navigate to System Settings. Mirrors working approach in supplement.py."""
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    clicked = False
    for s in ['div.nav-item:has-text("AcuHMI-1-7")',
              'div.nav-item-menu:has-text("AcuHMI-1-7")',
              '.nav-item-menu']:
        try:
            els = await page.query_selector_all(s)
            for el in els:
                txt = (await el.inner_text()).strip()
                if "AcuHMI" in txt or "acuhmi" in txt.lower():
                    await el.click()
                    clicked = True
                    break
            if clicked:
                break
        except Exception:
            pass

    if not clicked:
        try:
            await page.get_by_text("AcuHMI-1-7").last.click(timeout=3000)
            clicked = True
        except Exception:
            pass

    await settle(page, 10000)
    await page.wait_for_timeout(1500)   # extra settle for Vue rendering

    # Verify by reading left nav
    items = []
    try:
        els = await page.query_selector_all("li.left-nav-item")
        for el in els:
            t = (await el.inner_text()).strip()
            if t: items.append(t)
    except Exception:
        pass

    url = page.url
    log(f"  Current URL: {url}")
    log(f"  Left nav: {items}")

    in_sys = any("System" in t or "Templates" in t or "Protocols" in t
                 for t in items)
    if in_sys:
        log("  [OK] System Settings confirmed")
        return True
    log("  [WARN] System Settings nav not confirmed")
    return clicked   # return True if we at least clicked something

async def click_left_nav_item(page, label):
    try:
        els = await page.query_selector_all("li.left-nav-item")
        for el in els:
            t = (await el.inner_text()).strip()
            if t == label:
                await el.click()
                await settle(page, 5000)
                return True
    except Exception:
        pass
    try:
        await page.get_by_text(label, exact=True).first.click(timeout=3000)
        await settle(page, 5000)
        return True
    except Exception:
        pass
    return False

async def click_tab_by_text(page, label):
    """Click a tab element whose text exactly matches label. Returns True on success."""
    # Method 1: Playwright locators with different element types
    for tag in ["span", "a", "div", "li", "button"]:
        try:
            els = await page.query_selector_all(tag)
            for el in els:
                if not await el.is_visible():
                    continue
                txt = (await el.inner_text()).strip()
                if txt == label:
                    await el.click(timeout=2000)
                    await settle(page, 5000)
                    return True
        except Exception:
            pass

    # Method 2: JavaScript click
    try:
        result = await page.evaluate(f"""
            (function() {{
                const target = {repr(label)};
                const candidates = [];
                document.querySelectorAll('*').forEach(el => {{
                    if (el.children.length === 0 &&
                        el.textContent.trim() === target &&
                        el.offsetParent !== null) {{
                        candidates.push(el);
                    }}
                }});
                for (const el of candidates) {{
                    el.click();
                    return true;
                }}
                return false;
            }})()
        """)
        if result:
            await settle(page, 5000)
            return True
    except Exception:
        pass
    return False

# ── Part A: System Settings main tabs ────────────────────────────────────────
async def do_sys_settings_main_tabs(page):
    log("\n=== Part A-1: System Settings > System Settings tabs ===")
    folder = OUT_DIR / "SysSettings_01_System_Settings"

    if not await open_system_settings(page):
        log("  [ERR] Navigation failed")
        return

    if not await click_left_nav_item(page, "System Settings"):
        log("  [ERR] Cannot click System Settings left nav")
        return

    # Grab current URL as base
    base_url = page.url
    log(f"  Base URL: {base_url}")

    for tab in SYS_SETTINGS_TABS:
        log(f"  [TAB] {tab}")
        found = await click_tab_by_text(page, tab)
        if found:
            await shot(page, folder, f"tab_{clean(tab)}", f"System Settings > {tab}")
        else:
            log(f"  [MISS] {tab}")

# ── Part A-2: Maintenance tabs ────────────────────────────────────────────────
async def do_maintenance_tabs(page):
    log("\n=== Part A-2: Maintenance tabs ===")
    folder = OUT_DIR / "SysSettings_04_Maintenance"

    if not await open_system_settings(page):
        return
    if not await click_left_nav_item(page, "Maintenance"):
        log("  [ERR] Cannot navigate to Maintenance")
        return

    for tab in MAINTENANCE_TABS:
        log(f"  [TAB] {tab}")
        if DANGER.search(tab):
            log(f"  [SKIP] {tab}")
            continue
        found = await click_tab_by_text(page, tab)
        if found:
            await shot(page, folder, f"tab_{clean(tab)}", f"Maintenance > {tab}")
        else:
            log(f"  [MISS] {tab}")

# ── Part A-3: Other System Settings pages with potential tabs ─────────────────
async def do_other_sys_pages(page):
    log("\n=== Part A-3: Other System Settings pages ===")
    OTHER = [
        ("Templates",       OUT_DIR / "SysSettings_02_Templates"),
        ("Protocols",       OUT_DIR / "SysSettings_03_Protocols"),
        ("Diagnostics",     OUT_DIR / "SysSettings_05_Diagnostics"),
        ("User Management", OUT_DIR / "SysSettings_06_User_Management"),
        ("Firmware Update", OUT_DIR / "SysSettings_07_Firmware_Update"),
    ]
    for label, folder in OTHER:
        log(f"\n  {label}...")
        if not await open_system_settings(page):
            continue
        if not await click_left_nav_item(page, label):
            log(f"  [ERR] Cannot navigate to {label}")
            continue

        # Auto-detect tabs via bounding-box analysis
        tabs_js = await page.evaluate("""
            (function() {
                const result = [];
                document.querySelectorAll('*').forEach(el => {
                    if (el.children.length === 0 && el.offsetParent !== null) {
                        const txt = el.textContent.trim();
                        if (txt.length < 40 && txt.length > 2 && !txt.includes('\\n')) {
                            const r = el.getBoundingClientRect();
                            // Horizontal tab area: top portion of content, right of sidebar
                            if (r.top > 50 && r.top < 200 && r.left > 130 && r.height < 55) {
                                const cls = el.className || '';
                                const tag = el.tagName;
                                result.push({text: txt, cls: cls, tag: tag,
                                             top: r.top, left: r.left});
                            }
                        }
                    }
                });
                return result;
            })()
        """)
        if tabs_js:
            log(f"  Candidate tab elements: {[(t['text'], t['tag'], t['cls'][:30]) for t in tabs_js[:8]]}")
            seen = set()
            for item in tabs_js:
                t = item['text']
                if t in seen or DANGER.search(t): continue
                seen.add(t)
                found = await click_tab_by_text(page, t)
                if found:
                    await shot(page, folder, f"tab_{clean(t)}", f"{label} > {t}")

        # Probe Add/Create buttons for dialogs
        await probe_buttons_for_dialogs(page, folder, label)

async def probe_buttons_for_dialogs(page, folder, ctx, max_p=4):
    SAFE = re.compile(
        r"\b(add|create|new|edit|import|upload|设置|配置|新增|添加|编辑)\b",
        re.IGNORECASE,
    )
    probed = 0
    try:
        btns = await page.query_selector_all("button, .btn, [role='button']")
        for btn in btns:
            if probed >= max_p: break
            if not await btn.is_visible(): continue
            txt = (await btn.inner_text()).strip()
            if not txt or not SAFE.search(txt) or DANGER.search(txt): continue
            log(f"    Probing button: '{txt}'")
            try:
                await btn.click(timeout=2000)
                await page.wait_for_timeout(1000)
                for ds in ["[role='dialog']", ".ant-modal", ".el-dialog", ".modal"]:
                    d = await page.query_selector(ds)
                    if d and await d.is_visible():
                        fname = f"{nxt()}_dlg_{clean(txt)}.png"
                        fpath = folder / fname
                        folder.mkdir(parents=True, exist_ok=True)
                        await scroll_top(page)
                        await page.screenshot(path=str(fpath), full_page=False)
                        _log.append({"path": str(fpath),
                                     "desc": f"{ctx} - Dialog via '{txt}'"})
                        log(f"    [SHOT] {fname}")
                        await close_dialog(page)
                        probed += 1
                        break
            except Exception:
                pass
    except Exception:
        pass

# ── Part B: Add Device dialogs ────────────────────────────────────────────────
async def do_device_dialogs(page):
    log("\n=== Part B: Add Device Dialogs ===")
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    # Ensure Devices context
    try:
        els = await page.query_selector_all("div.nav-item-menu")
        for el in els:
            if (await el.inner_text()).strip() == "Devices":
                await el.click(); break
    except Exception: pass
    await settle(page)

    for lbl, fnum in [("Physical Devices", 2), ("Virtual Devices", 3)]:
        folder = OUT_DIR / f"{fnum:02d}_{clean(lbl)}"
        log(f"\n  {lbl}...")

        if not await click_left_nav_item(page, lbl):
            log(f"  [ERR] Cannot nav to {lbl}")
            continue

        # Find "Add Device" button — try multiple text patterns
        found_btn = None
        for pat in ['button:has-text("Add Device")', 'button:has-text("Add")',
                    '.el-button:has-text("Add")', '[class*="add"]:visible']:
            try:
                btns = await page.query_selector_all(pat)
                for b in btns:
                    if await b.is_visible():
                        txt = (await b.inner_text()).strip()
                        log(f"    Found button: '{txt}'")
                        if "Add" in txt and not DANGER.search(txt):
                            found_btn = b; break
                if found_btn: break
            except Exception: pass

        if found_btn:
            try:
                await found_btn.click(timeout=2000)
                await page.wait_for_timeout(1200)
                for ds in ["[role='dialog']", ".ant-modal", ".el-dialog", ".modal"]:
                    d = await page.query_selector(ds)
                    if d and await d.is_visible():
                        fname = f"{nxt()}_dlg_Add_Device.png"
                        fpath = folder / fname
                        folder.mkdir(parents=True, exist_ok=True)
                        await page.screenshot(path=str(fpath), full_page=False)
                        _log.append({"path": str(fpath),
                                     "desc": f"{lbl} - Add Device Dialog"})
                        log(f"  [SHOT] {fname}")
                        await close_dialog(page)
                        break
            except Exception as e:
                log(f"  [ERR] {e}")
        else:
            log(f"  [WARN] No Add Device button found for {lbl}")

def write_report():
    rp = OUT_DIR / "tabs_index.md"
    lines = [
        "# AcuHMI-1-7 Tab Screenshots",
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
    log("  AcuHMI-1-7 Pass 3 – Tabs & Dialogs")
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

        await do_sys_settings_main_tabs(page)
        await do_maintenance_tabs(page)
        await do_other_sys_pages(page)
        await do_device_dialogs(page)

        await browser.close()

    write_report()
    log("\n" + "=" * 60)
    log(f"  Pass 3 done! {_n} screenshots.")
    log("=" * 60)

if __name__ == "__main__":
    asyncio.run(main())

