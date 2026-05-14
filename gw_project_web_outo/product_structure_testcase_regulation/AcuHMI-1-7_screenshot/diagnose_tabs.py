# -*- coding: utf-8 -*-
"""Find the CSS selectors for tabs in AcuHMI-1-7 System Settings"""
import asyncio, sys
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_URL = "https://192.168.2.199"
USERNAME = "admin"
PASSWORD = "Admin@110001"

async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True, args=["--ignore-certificate-errors"]
        )
        ctx = await browser.new_context(ignore_https_errors=True,
                                        viewport={"width": 1920, "height": 1080})
        page = await ctx.new_page()

        # Login
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
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception: pass
        await page.wait_for_timeout(1500)

        # Open System Settings
        for s in ['div.nav-item-menu', 'div.nav-item']:
            try:
                els = await page.query_selector_all(s)
                for el in els:
                    txt = (await el.inner_text()).strip()
                    if "AcuHMI" in txt:
                        await el.click()
                        await page.wait_for_timeout(1000)
                        break
            except Exception: pass

        await page.wait_for_load_state("networkidle", timeout=10000)
        await page.wait_for_timeout(1000)

        print("\n--- Looking for tab-like elements ---")
        # Inspect horizontal nav/tab items visible on System Settings page
        for sel in [
            "a", "span", "li", "div", "button"
        ]:
            try:
                els = await page.query_selector_all(sel)
                matches = []
                for el in els:
                    if not await el.is_visible(): continue
                    txt = (await el.inner_text()).strip()
                    # Looking for known tab labels
                    if txt in ["Date & Time", "Network", "Access Control", "Email",
                               "Alarm Notification", "Certificate Management",
                               "Configuration Management", "Remote Access",
                               "System Status", "Event Log"]:
                        cls = await el.get_attribute("class") or ""
                        tag = await el.evaluate("e => e.tagName")
                        parent_cls = await el.evaluate(
                            "e => e.parentElement ? e.parentElement.className : ''"
                        )
                        matches.append(f"  <{tag}> class={cls!r:60s}  parent_cls={parent_cls[:60]!r}  text={txt!r}")
                if matches:
                    print(f"\n  Matches via <{sel}>:")
                    for m in matches:
                        print(m)
                    break
            except Exception: pass

        print("\n--- All visible short-text items (possible tabs) ---")
        for sel in ["a", "li", "span", "div"]:
            try:
                els = await page.query_selector_all(sel)
                for el in els:
                    if not await el.is_visible(): continue
                    txt = (await el.inner_text()).strip()
                    if not txt or len(txt) > 35 or "\n" in txt: continue
                    cls = await el.get_attribute("class") or ""
                    if any(w in cls.lower() for w in ["tab", "nav", "item", "link", "menu"]):
                        tag = await el.evaluate("e => e.tagName")
                        print(f"  <{tag}> class={cls[:70]!r}  text={txt!r}")
            except Exception: pass

        await browser.close()

asyncio.run(main())
