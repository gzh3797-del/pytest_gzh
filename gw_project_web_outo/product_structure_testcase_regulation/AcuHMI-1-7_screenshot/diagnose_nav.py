# -*- coding: utf-8 -*-
"""Diagnose the nav DOM structure of AcuHMI-1-7"""
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
        await page.goto(BASE_URL, wait_until="networkidle", timeout=30000)

        # Login
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
                if b and await b.is_visible():
                    await b.click(); break
            except Exception: pass
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception: pass
        await page.wait_for_timeout(1500)

        print("\n--- Sidebar / nav candidates ---")
        candidates = [
            "nav", "aside", ".sidebar", ".side-bar", ".left-menu",
            ".nav-menu", ".menu", "[class*='sidebar']", "[class*='nav']",
            "[class*='menu']", "[id*='nav']", "[id*='menu']", "[id*='sidebar']",
        ]
        for sel in candidates:
            try:
                els = await page.query_selector_all(sel)
                visible = [e for e in els if await e.is_visible()]
                if visible:
                    for el in visible[:2]:
                        txt = (await el.inner_text()).strip()[:120]
                        cls = await el.get_attribute("class") or ""
                        tag = await el.evaluate("e => e.tagName")
                        print(f"  sel={sel!r:35s}  tag={tag}  class={cls[:60]!r}  text={txt!r}")
            except Exception: pass

        print("\n--- Visible clickable items in first 8 tag matches ---")
        for tag in ["li", "div", "span", "a"]:
            els = await page.query_selector_all(tag)
            count = 0
            for el in els:
                try:
                    if count >= 8: break
                    if not await el.is_visible(): continue
                    txt = (await el.inner_text()).strip()
                    if not txt or len(txt) > 40 or "\n" in txt: continue
                    cls = await el.get_attribute("class") or ""
                    # Only show plausibly nav-like items
                    nav_words = any(w in cls.lower() for w in
                                    ["menu","nav","sidebar","item","link"])
                    if nav_words or tag == "a":
                        print(f"  <{tag}> class={cls[:60]!r}  text={txt!r}")
                        count += 1
                except Exception: pass

        print("\n--- Full sidebar HTML (first 3000 chars) ---")
        for sel in ["aside", ".sidebar", "[class*='sidebar']", "[class*='nav']"]:
            try:
                el = await page.query_selector(sel)
                if el and await el.is_visible():
                    html = await el.inner_html()
                    print(f"  [{sel}]\n{html[:3000]}\n")
                    break
            except Exception: pass

        await browser.close()

asyncio.run(main())
