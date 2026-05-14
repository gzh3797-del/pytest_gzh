# -*- coding: utf-8 -*-
"""
AcuHMI-1-7 Supplementary Screenshot Pass
1. System Settings section (triggered by AcuHMI-1-7 top-nav)
2. Add Device dialogs for Physical / Virtual / Web Devices
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
    r"删除|恢复出厂|清除|格式化|重置",
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
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(300)

async def shot(page, folder, name, desc=""):
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}.png"
    fpath = folder / fname
    await scroll_top(page)
    await page.wait_for_timeout(400)
    try:
        await page.screenshot(path=str(fpath), full_page=False)
        _log.append({"path": str(fpath), "desc": desc or name})
        log(f"  [SHOT] {fname}  ({desc or name})")
        return fpath
    except Exception as e:
        log(f"  [WARN] shot failed ({name}): {e}")
        return None

async def shot_fp(page, folder, name, desc=""):
    """Full-page variant."""
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}_fp.png"
    fpath = folder / fname
    await scroll_top(page)
    await page.wait_for_timeout(400)
    try:
        await page.screenshot(path=str(fpath), full_page=True)
        _log.append({"path": str(fpath), "desc": (desc or name) + " (full)"})
        log(f"  [SHOT] {fname}  ({desc or name} full)")
        return fpath
    except Exception as e:
        log(f"  [WARN] shot_fp failed ({name}): {e}")
        return None

async def settle(page, ms=8000):
    try:
        await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout:
        pass
    await page.wait_for_timeout(800)

async def login(page):
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
                return True
        except Exception: pass
    try:
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(500)
    except Exception: pass
    return False

# ---- Part A: System Settings ------------------------------------------------
# Left sidebar selectors inside System Settings
SYS_LEFT_SEL = "li.left-nav-item"

async def do_system_settings(page):
    log("\n=== Part A: System Settings ===")

    # Navigate to System Settings via AcuHMI-1-7 top button
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    # Click the AcuHMI-1-7 button
    clicked = False
    for s in ['div.nav-item:has-text("AcuHMI-1-7")',
              'div.nav-item-menu:has-text("AcuHMI-1-7")',
              '.nav-item-menu']:
        try:
            els = await page.query_selector_all(s)
            for el in els:
                txt = (await el.inner_text()).strip()
                if "AcuHMI" in txt or "acuhmi" in txt.lower():
                    await el.click(); clicked = True; break
            if clicked: break
        except Exception: pass

    if not clicked:
        try:
            await page.get_by_text("AcuHMI-1-7").last.click(timeout=3000)
            clicked = True
        except Exception: pass

    if not clicked:
        log("  [ERR] Cannot open System Settings")
        return

    await settle(page)
    log("  Opened System Settings")

    # Read the left sidebar in System Settings
    await page.wait_for_timeout(1000)
    left_items = []
    try:
        els = await page.query_selector_all(SYS_LEFT_SEL)
        for i, el in enumerate(els):
            try:
                lbl = (await el.inner_text()).strip()
                if lbl:
                    left_items.append({"label": lbl, "idx": i})
            except Exception: pass
    except Exception: pass

    log(f"  System Settings left nav: {[x['label'] for x in left_items]}")

    for num, item in enumerate(left_items, start=1):
        lbl    = item["label"]
        folder = OUT_DIR / f"SysSettings_{num:02d}_{clean(lbl)}"
        log(f"\n  --- SysSet [{num:02d}] {lbl} ---")

        if DANGER.search(lbl):
            log(f"    [SKIP] dangerous: {lbl}")
            continue

        # Click this menu item
        try:
            els = await page.query_selector_all(SYS_LEFT_SEL)
            if item["idx"] < len(els):
                await els[item["idx"]].click()
            else:
                raise IndexError("stale")
        except Exception:
            try:
                await page.get_by_text(lbl, exact=True).first.click(timeout=3000)
            except Exception:
                log(f"    [ERR] cannot click: {lbl}")
                continue

        await settle(page)
        await shot(page, folder, "main", f"SysSettings > {lbl}")
        await shot_fp(page, folder, "main", f"SysSettings > {lbl}")

        # Screenshot all tabs on this page
        await do_tabs(page, folder, lbl)

        # Screenshot Add/Create dialogs
        await probe_add_buttons(page, folder, lbl)

    log("\n  System Settings done")

async def do_tabs(page, folder, parent):
    """Find and screenshot all horizontal tabs on the current page."""
    TAB_SELS = [
        ".tab-list .tab-item", ".tabs .tab", "[role='tab']",
        ".el-tabs__item", ".ant-tabs-tab",
        # This product seems to use its own tab bar—try these too
        ".tab-bar a", ".tab-bar span", ".tab-bar div",
        ".tab-wrap .tab", ".header-tabs .tab",
        # Generic clickable tab-like elements in header area
        "header .item", ".page-header .item",
    ]
    tabs = []
    seen = set()

    for sel in TAB_SELS:
        try:
            els = await page.query_selector_all(sel)
            for el in els:
                if not await el.is_visible(): continue
                lbl = (await el.inner_text()).strip()
                if not lbl or lbl in seen or len(lbl) > 50: continue
                seen.add(lbl)
                tabs.append({"el": el, "label": lbl})
            if len(tabs) >= 2:
                log(f"    Found {len(tabs)} tabs via '{sel}'")
                break
        except Exception: pass

    if len(tabs) < 2:
        # Try to find tabs by inspecting visible headings inside a tab container
        return

    for tab in tabs:
        lbl = tab["label"]
        if DANGER.search(lbl): continue
        log(f"      [TAB] {lbl}")
        try:
            await tab["el"].click(timeout=2000)
            await settle(page, 5000)
            await shot(page, folder, f"tab_{clean(lbl)}", f"{parent} > {lbl}")
            await shot_fp(page, folder, f"tab_{clean(lbl)}", f"{parent} > {lbl}")
        except Exception as e:
            log(f"      [ERR] tab '{lbl}': {e}")

async def probe_add_buttons(page, folder, ctx, max_probes=5):
    SAFE = re.compile(
        r"\b(add|create|new|edit|config|configure|detail|view|import|"
        r"设置|配置|新增|添加|编辑|详情)\b",
        re.IGNORECASE,
    )
    probed = 0
    try:
        btns = await page.query_selector_all("button, .btn, [role='button'], a.btn")
        for btn in btns:
            if probed >= max_probes: break
            try:
                if not await btn.is_visible(): continue
                txt = (await btn.inner_text()).strip()
                if not txt or not SAFE.search(txt) or DANGER.search(txt): continue

                await btn.click(timeout=2000)
                await page.wait_for_timeout(800)

                dialog_found = False
                for ds in ["[role='dialog']", ".ant-modal", ".el-dialog",
                           ".el-message-box", ".modal"]:
                    d = await page.query_selector(ds)
                    if d and await d.is_visible():
                        await scroll_top(page)
                        fname = f"{nxt()}_dlg_{clean(txt)}.png"
                        fpath = folder / fname
                        await page.screenshot(path=str(fpath), full_page=False)
                        _log.append({"path": str(fpath), "desc": f"Dialog via '{txt}' ({ctx})"})
                        log(f"    [SHOT] {fname}  (dialog via '{txt}')")
                        await close_dialog(page)
                        probed += 1
                        dialog_found = True
                        break

                if not dialog_found:
                    # Button may have navigated — go back
                    await page.go_back()
                    await settle(page, 5000)
            except Exception:
                pass
    except Exception:
        pass

# ---- Part B: Add Device dialogs (Physical / Virtual / Web) ------------------
LEFT_NAV_SEL = "li.left-nav-item"

DEVICE_PAGES = [
    {"label": "Physical Devices", "nav_idx": 1},
    {"label": "Virtual Devices",  "nav_idx": 2},
    {"label": "Web Devices",      "nav_idx": 3},
]

async def do_add_device_dialogs(page):
    log("\n=== Part B: Add Device Dialogs ===")

    # Back to the Devices section
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    # Click "Devices" top nav to ensure we are in Devices context
    try:
        els = await page.query_selector_all("div.nav-item-menu")
        for el in els:
            txt = (await el.inner_text()).strip()
            if txt == "Devices":
                await el.click(); break
    except Exception: pass
    await settle(page)

    for dp in DEVICE_PAGES:
        lbl    = dp["label"]
        folder = OUT_DIR / f"{dp['nav_idx']:02d}_{clean(lbl)}"
        log(f"\n  --- {lbl} dialogs ---")

        # Click the left nav item
        try:
            els = await page.query_selector_all(LEFT_NAV_SEL)
            # Find by text match (safer than index when pages change)
            clicked = False
            for el in els:
                t = (await el.inner_text()).strip()
                if t == lbl:
                    await el.click(); clicked = True; break
            if not clicked:
                raise ValueError("not found by text")
        except Exception:
            try:
                await page.get_by_text(lbl, exact=True).first.click(timeout=3000)
            except Exception:
                log(f"  [ERR] cannot nav to {lbl}")
                continue

        await settle(page)

        # Probe Add/Edit buttons
        await probe_add_buttons(page, folder, lbl, max_probes=6)

    log("\n  Add Device dialogs done")

# ---- Part C: Alarm / Data Log additional probing ----------------------------
async def do_alarm_datalog(page):
    log("\n=== Part C: Alarm & Data Log detailed probing ===")

    for item_label in ["Alarm", "Data Log"]:
        folder = OUT_DIR / f"{['Alarm','Data Log'].index(item_label)+5:02d}_{clean(item_label)}"

        # Navigate back to Devices context first
        await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
        await settle(page)

        try:
            els = await page.query_selector_all(LEFT_NAV_SEL)
            for el in els:
                t = (await el.inner_text()).strip()
                if t == item_label:
                    await el.click(); break
        except Exception:
            try:
                await page.get_by_text(item_label, exact=True).first.click(timeout=3000)
            except Exception:
                continue

        await settle(page)
        await do_tabs(page, folder, item_label)
        await probe_add_buttons(page, folder, item_label)

    log("\n  Alarm & Data Log done")

# ---- Write supplement report ------------------------------------------------
def write_report():
    rp = OUT_DIR / "supplement_index.md"
    lines = [
        "# AcuHMI-1-7 Supplement Screenshots",
        "",
        f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}",
        f"New screenshots: {_n}",
        "",
        "| # | File | Description |",
        "|---|------|-------------|",
    ]
    for i, e in enumerate(_log, 1):
        try:
            rel = Path(e["path"]).relative_to(OUT_DIR)
        except Exception:
            rel = e.get("path", "")
        lines.append(f"| {i} | `{rel}` | {e['desc']} |")
    rp.write_text("\n".join(lines), encoding="utf-8")
    log(f"  [OK] Report: {rp}")

# ---- Main -------------------------------------------------------------------
async def main():
    log("=" * 60)
    log("  AcuHMI-1-7 Supplementary Screenshots")
    log(f"  Output: {OUT_DIR}")
    log("=" * 60)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", f"--window-size={W},{H}"],
        )
        ctx = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": W, "height": H},
        )
        page = await ctx.new_page()

        await login(page)

        # A: System Settings (full exploration)
        await do_system_settings(page)

        # B: Add Device dialogs
        await do_add_device_dialogs(page)

        # C: Alarm / Data Log tabs
        await do_alarm_datalog(page)

        await browser.close()

    write_report()
    log("\n" + "=" * 60)
    log(f"  Supplement done! {_n} additional screenshots.")
    log("=" * 60)

if __name__ == "__main__":
    asyncio.run(main())

