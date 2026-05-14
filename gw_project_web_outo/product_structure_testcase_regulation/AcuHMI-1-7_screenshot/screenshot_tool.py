# -*- coding: utf-8 -*-
"""
AcuHMI-1-7 Gateway Screenshot Automation
Selectors confirmed via DOM inspection.
Full-viewport screenshots (no partial / obstructed captures).
"""
import asyncio, json, re, sys, datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------- Config -----------------------------------------------------------
BASE_URL = "https://192.168.2.199"
USERNAME = "admin"
PASSWORD = "Admin@110001"
OUT_DIR  = Path(r"C:\knowledge_base\AcuHMI-1-7_screenshot")
W, H     = 1920, 1080          # full-HD viewport

# Destructive actions — screenshot only, never click
DANGER = re.compile(
    r"delete|remove|restore|reset.factory|factory.reset|clear.all|erase|format|"
    r"删除|恢复出厂|清除|格式化|重置",
    re.IGNORECASE,
)

# ---------- State ------------------------------------------------------------
_n         = 0
_log       = []
_hierarchy = {}      # {"Module": {"folder": ..., "subs": [...]}}

# ---------- Helpers ----------------------------------------------------------
def clean(name):
    name = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", name)
    return re.sub(r"\s+", "_", name.strip())[:60] or "unnamed"

def nxt():
    global _n; _n += 1
    return f"{_n:04d}"

def log(msg):
    print(msg, flush=True)

async def scroll_top(page):
    """Scroll to top so nothing is cut off."""
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(300)

async def close_any_open_dropdown(page):
    """Press Escape to close dropdowns/overlays before screenshot."""
    try:
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(300)
    except Exception:
        pass

async def shot(page, folder, name, desc=""):
    """Full-viewport screenshot (1920×1080). Scrolls to top first."""
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}.png"
    fpath = folder / fname
    await scroll_top(page)
    await close_any_open_dropdown(page)
    await page.wait_for_timeout(400)          # let animations settle
    try:
        # full_page=False → captures exactly the 1920×1080 viewport, no clipping by overlays
        await page.screenshot(path=str(fpath), full_page=False)
        _log.append({"path": str(fpath), "desc": desc or name})
        log(f"  [SHOT] {fname}  ({desc or name})")
        return fpath
    except Exception as e:
        log(f"  [WARN] screenshot failed ({name}): {e}")
        return None

async def shot_fullpage(page, folder, name, desc=""):
    """Full-page screenshot for scrollable content."""
    folder.mkdir(parents=True, exist_ok=True)
    fname = f"{nxt()}_{clean(name)}_fullpage.png"
    fpath = folder / fname
    await scroll_top(page)
    await page.wait_for_timeout(400)
    try:
        await page.screenshot(path=str(fpath), full_page=True)
        _log.append({"path": str(fpath), "desc": (desc or name) + " (full page)"})
        log(f"  [SHOT] {fname}  ({desc or name} - full page)")
        return fpath
    except Exception as e:
        log(f"  [WARN] full-page screenshot failed ({name}): {e}")
        return None

async def settle(page, ms=8000):
    try:
        await page.wait_for_load_state("networkidle", timeout=ms)
    except PWTimeout:
        pass
    await page.wait_for_timeout(800)

async def shot_dialogs(page, folder, ctx):
    """Capture visible modals/dialogs."""
    for sel in ["[role='dialog']", ".ant-modal", ".el-dialog",
                ".el-message-box", ".modal", ".popup", ".overlay"]:
        try:
            for i, el in enumerate(await page.query_selector_all(sel)):
                if await el.is_visible():
                    await page.wait_for_timeout(300)
                    await page.screenshot(
                        path=str(folder / f"{nxt()}_dialog_{clean(ctx)}_{i}.png"),
                        full_page=False,
                    )
                    _log.append({"path": str(folder / f"dialog_{ctx}_{i}.png"),
                                 "desc": f"Dialog - {ctx}"})
                    log(f"  [SHOT] dialog - {ctx}")
        except Exception:
            pass

# ---------- Phase 1 : Login --------------------------------------------------
async def do_login(page):
    log("\n=== Phase 1: Login ===")
    d = OUT_DIR / "00_Login"
    await page.goto(BASE_URL, wait_until="networkidle", timeout=30000)
    await shot(page, d, "01_login_page", "Login Page")

    u_sel = None
    for s in ['input[name="username"]', 'input[name="user"]',
              '#username', '#user', 'input[type="text"]',
              'input[placeholder*="user" i]']:
        try:
            await page.wait_for_selector(s, timeout=2000, state="visible")
            u_sel = s; break
        except PWTimeout:
            pass

    p_sel = None
    try:
        await page.wait_for_selector('input[type="password"]', timeout=3000, state="visible")
        p_sel = 'input[type="password"]'
    except PWTimeout:
        pass

    if not (u_sel and p_sel):
        log("  [ERR] Cannot find login fields")
        return False

    await page.fill(u_sel, USERNAME)
    await page.fill(p_sel, PASSWORD)
    await shot(page, d, "02_login_filled", "Login Form Filled")

    submitted = False
    for s in ['button[type="submit"]', 'input[type="submit"]',
              'button:has-text("Login")', 'button:has-text("Sign In")',
              'button:has-text("Log in")', 'button:has-text("登录")',
              '.login-btn', '#login-btn', 'button']:
        try:
            b = await page.query_selector(s)
            if b and await b.is_visible():
                await b.click(); submitted = True; break
        except Exception:
            pass
    if not submitted:
        await page.keyboard.press("Enter")

    await settle(page, 15000)
    await shot(page, d, "03_home_dashboard", "Home - Dashboard")
    await shot_fullpage(page, d, "03_home_dashboard", "Home - Dashboard")
    log("  [OK] Login done")
    return True

# ---------- Phase 2 : Left sidebar pages ------------------------------------
LEFT_NAV_SEL = "li.left-nav-item"

async def do_left_nav(page):
    log("\n=== Phase 2: Left Sidebar Pages ===")

    # Read all left-nav labels first
    items = []
    try:
        els = await page.query_selector_all(LEFT_NAV_SEL)
        for i, el in enumerate(els):
            try:
                lbl = (await el.inner_text()).strip()
                if lbl:
                    items.append({"label": lbl, "idx": i})
            except Exception:
                pass
    except Exception:
        pass

    log(f"  Found {len(items)} left-nav items: {[x['label'] for x in items]}")

    for num, item in enumerate(items, start=1):
        lbl    = item["label"]
        folder = OUT_DIR / f"{num:02d}_{clean(lbl)}"
        log(f"\n  --- [{num:02d}] {lbl} ---")

        if DANGER.search(lbl):
            log(f"    [SKIP] dangerous: {lbl}")
            _hierarchy[lbl] = {"skipped": True}
            continue

        # Re-fetch and click
        try:
            els = await page.query_selector_all(LEFT_NAV_SEL)
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
        await shot(page, folder, "main", f"{lbl} - Main View")
        await shot_fullpage(page, folder, "main", f"{lbl} - Main View")
        await shot_dialogs(page, folder, lbl)
        _hierarchy[lbl] = {"folder": str(folder), "subs": []}

        # Probe Add/Edit buttons for dialogs
        await probe_buttons(page, folder, lbl)

        # Look for sub-tabs / secondary nav on this page
        await do_page_tabs(page, folder, lbl)

    return items

# ---------- Phase 3 : Top-nav menus (About, Devices, AcuHMI-1-7) -------------
TOP_MENU_SEL = "div.nav-item-menu"

async def do_top_nav(page, left_nav_items):
    log("\n=== Phase 3: Top Nav Menus ===")

    # Start fresh at home
    await page.goto(BASE_URL, wait_until="networkidle", timeout=20000)
    await settle(page)

    menus = []
    try:
        els = await page.query_selector_all(TOP_MENU_SEL)
        for i, el in enumerate(els):
            lbl = (await el.inner_text()).strip()
            if lbl:
                menus.append({"label": lbl, "idx": i})
    except Exception:
        pass

    log(f"  Found {len(menus)} top-nav menus: {[x['label'] for x in menus]}")

    base_folder_num = len(left_nav_items) + 1

    for num, menu in enumerate(menus, start=base_folder_num):
        lbl    = menu["label"]
        folder = OUT_DIR / f"{num:02d}_TopNav_{clean(lbl)}"
        log(f"\n  --- Top [{num:02d}] {lbl} ---")

        if DANGER.search(lbl):
            log(f"    [SKIP] dangerous: {lbl}")
            continue
        if lbl.lower() == "logout":
            log(f"    [SKIP] Logout - not clicking")
            continue

        try:
            els = await page.query_selector_all(TOP_MENU_SEL)
            if menu["idx"] < len(els):
                await els[menu["idx"]].click()
            else:
                raise IndexError("stale")
        except Exception:
            try:
                await page.get_by_text(lbl, exact=True).first.click(timeout=3000)
            except Exception:
                log(f"    [ERR] cannot click: {lbl}")
                continue

        await page.wait_for_timeout(1000)
        await shot(page, folder, "main", f"Top Nav - {lbl}")

        # Screenshot dropdown items that appeared
        await do_dropdown_items(page, folder, lbl)

        # Press Escape to close any open dropdown
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(500)

        _hierarchy[f"[TopNav] {lbl}"] = {"folder": str(folder), "subs": []}

# ---------- In-page tabs / secondary nav ------------------------------------
async def do_page_tabs(page, folder, parent):
    """Click secondary tabs on a page and screenshot each."""
    TAB_SELS = [
        ".tabs .tab-item", ".tab-list .tab", "[role='tab']",
        ".el-tabs__item", ".ant-tabs-tab",
        "ul.tab-nav > li", ".page-tabs a",
    ]
    tabs = []
    seen = set()
    for sel in TAB_SELS:
        try:
            els = await page.query_selector_all(sel)
            for el in els:
                if not await el.is_visible():
                    continue
                lbl = (await el.inner_text()).strip()
                if not lbl or lbl in seen or len(lbl) > 40:
                    continue
                seen.add(lbl)
                tabs.append({"el": el, "label": lbl})
            if tabs:
                break
        except Exception:
            pass

    if len(tabs) <= 1:
        return

    log(f"    Found {len(tabs)} tabs under '{parent}'")
    for tab in tabs:
        lbl = tab["label"]
        if DANGER.search(lbl):
            log(f"      [SKIP] dangerous tab: {lbl}")
            continue
        try:
            await tab["el"].click(timeout=2000)
            await settle(page, 5000)
            await shot(page, folder, f"tab_{clean(lbl)}", f"{parent} > Tab: {lbl}")
            await shot_fullpage(page, folder, f"tab_{clean(lbl)}", f"{parent} > Tab: {lbl}")
            await shot_dialogs(page, folder, f"{parent}_tab_{lbl}")
            _hierarchy.get(parent, {}).get("subs", []).append(f"Tab: {lbl}")
        except Exception as e:
            log(f"      [ERR] tab '{lbl}': {e}")

# ---------- Dropdown items ---------------------------------------------------
async def do_dropdown_items(page, folder, parent):
    """Screenshot dropdown menu items that appeared after clicking a top-nav item."""
    DROP_SELS = [
        ".dropdown-menu a", ".dropdown-menu li", ".ant-dropdown-menu-item",
        ".el-dropdown-menu__item", ".popover-content a",
        "[class*='dropdown'] li", "[class*='dropdown'] a",
    ]
    items = []
    seen  = set()
    for sel in DROP_SELS:
        try:
            els = await page.query_selector_all(sel)
            for el in els:
                if not await el.is_visible():
                    continue
                lbl = (await el.inner_text()).strip()
                if not lbl or lbl in seen or len(lbl) > 40:
                    continue
                seen.add(lbl)
                items.append({"el": el, "label": lbl})
            if items:
                break
        except Exception:
            pass

    if not items:
        return

    log(f"    Dropdown has {len(items)} items under '{parent}'")
    for item in items:
        lbl = item["label"]
        log(f"      + {lbl}")
        if DANGER.search(lbl):
            log(f"        [SKIP] dangerous: {lbl}")
            await shot(page, folder, f"skip_{clean(lbl)}", f"[NOT CLICKED] {lbl}")
            continue
        try:
            await item["el"].click(timeout=2000)
            await settle(page, 5000)
            await shot(page, folder, f"dd_{clean(lbl)}", f"{parent} > {lbl}")
            await shot_fullpage(page, folder, f"dd_{clean(lbl)}", f"{parent} > {lbl}")
            await shot_dialogs(page, folder, f"{parent}_{lbl}")
            _hierarchy.get(f"[TopNav] {parent}", {}).get("subs", []).append(lbl)
        except Exception as e:
            log(f"        [ERR] {lbl}: {e}")
        # Re-open the dropdown for the next item
        try:
            els = await page.query_selector_all(TOP_MENU_SEL)
            for el in els:
                t = (await el.inner_text()).strip()
                if t == parent:
                    await el.click()
                    await page.wait_for_timeout(600)
                    break
        except Exception:
            pass

# ---------- Probe Add/Edit buttons for dialogs ------------------------------
async def probe_buttons(page, folder, ctx):
    SAFE = re.compile(
        r"\b(add|create|new|edit|config|configure|detail|view|import|"
        r"设置|配置|新增|添加|编辑|详情)\b",
        re.IGNORECASE,
    )
    probed = 0
    try:
        btns = await page.query_selector_all("button, .btn, [role='button']")
        for btn in btns:
            if probed >= 4:
                break
            try:
                if not await btn.is_visible():
                    continue
                txt = (await btn.inner_text()).strip()
                if not txt or not SAFE.search(txt) or DANGER.search(txt):
                    continue
                await btn.click(timeout=2000)
                await page.wait_for_timeout(800)
                for ds in ["[role='dialog']", ".ant-modal", ".el-dialog", ".modal"]:
                    d = await page.query_selector(ds)
                    if d and await d.is_visible():
                        await page.screenshot(
                            path=str(folder / f"{nxt()}_dlg_{clean(txt)}.png"),
                            full_page=False,
                        )
                        _log.append({"path": "", "desc": f"Dialog via '{txt}' ({ctx})"})
                        log(f"    [SHOT] dialog via button '{txt}'")
                        for cs in [".ant-modal-close", ".el-dialog__headerbtn",
                                   'button:has-text("Cancel")', 'button:has-text("取消")',
                                   'button:has-text("Close")', "[aria-label='Close']"]:
                            try:
                                c = await page.query_selector(cs)
                                if c and await c.is_visible():
                                    await c.click(timeout=1500); break
                            except Exception:
                                pass
                        await page.wait_for_timeout(500)
                        probed += 1
                        break
            except Exception:
                pass
    except Exception:
        pass

# ---------- Phase 4 : Report -------------------------------------------------
def write_report():
    log("\n=== Phase 4: Writing Report ===")
    rp = OUT_DIR / "page_structure.md"
    lines = [
        "# AcuHMI-1-7 Page Structure",
        "",
        f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}",
        f"URL:  {BASE_URL}",
        f"Total screenshots: {_n}",
        "",
        "---",
        "## Page Hierarchy",
        "",
    ]
    for mod, info in _hierarchy.items():
        if info.get("skipped"):
            lines.append(f"- ~~{mod}~~ (skipped)")
        else:
            lines.append(f"- **{mod}**")
            for sp in info.get("subs", []):
                lines.append(f"  - {sp}")
    lines += [
        "",
        "---",
        "## Screenshot Index",
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

# ---------- Main -------------------------------------------------------------
async def main():
    log("=" * 60)
    log("  AcuHMI-1-7 Screenshot Tool  (full-viewport mode)")
    log(f"  Target  : {BASE_URL}")
    log(f"  Output  : {OUT_DIR}")
    log(f"  Viewport: {W}x{H}")
    log("=" * 60)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors",
                  f"--window-size={W},{H}",
                  "--start-maximized"],
        )
        ctx = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": W, "height": H},
        )
        page = await ctx.new_page()

        # Phase 1 – Login
        if not await do_login(page):
            await browser.close()
            return

        # Phase 2 – Left sidebar
        left_items = await do_left_nav(page)

        # Phase 3 – Top nav menus
        await do_top_nav(page, left_items)

        await browser.close()

    # Phase 4 – Report
    write_report()

    log("\n" + "=" * 60)
    log(f"  Done!  {_n} screenshots captured.")
    log(f"  Report : {OUT_DIR / 'page_structure.md'}")
    log(f"  Folder : {OUT_DIR}")
    log("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())

