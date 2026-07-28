# -*- coding: utf-8 -*-
"""
parameter_settings/conftest.py
parameter_settings 模块专用 fixtures

- app_page : session 级，整个会话只启动一次浏览器并登录，
             供 acurev4100 / acuvim_iiw / acurev2100 等子模块复用。
- screenshot_on_fail : 用例失败时自动截图，保存到 logs/ 并附加到 Allure 报告。

配置统一从 projects/RPP/settings.py 读取（BASE_URL / GATEWAY_WEB_* 指向 RPP 网关）。
"""
import sys
from pathlib import Path

import allure
import pytest
from playwright.sync_api import Browser, Page

# ── 确保仓库根目录在 sys.path，可 import projects.RPP.settings ────────────────
_REPO_ROOT = Path(__file__).resolve().parents[4]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import projects.RPP.settings as _cfg

_LOGS_DIR = _REPO_ROOT / "logs"
_LOGS_DIR.mkdir(parents=True, exist_ok=True)


# ── app_page：session 级已登录页面 ────────────────────────────────────────────
@pytest.fixture(scope="package")
def app_page(browser: Browser) -> Page:
    """复用 pytest-playwright 提供的 browser（启动参数由外层 conftest 配置），
    自行创建 context + page 并完成登录，整个 session 共享同一个已登录页面。"""
    ctx = browser.new_context(
        ignore_https_errors=True,
        viewport={"width": 1920, "height": 1080},
    )
    pg = ctx.new_page()
    _login(pg)
    yield pg
    ctx.close()


def _login(page: Page) -> None:
    url  = _cfg.GATEWAY_WEB_URL
    user = _cfg.GATEWAY_WEB_USER
    pwd  = _cfg.GATEWAY_WEB_PASS
    print(f"\n  [LOGIN] {url}  user={user}", flush=True)

    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)

    page.locator(
        "input[type='text'], input[name*='user' i], input[id*='user' i]"
    ).first.fill(user)
    page.locator("input[type='password']").first.fill(pwd)
    page.locator("button:has-text('Sign In'), button[type='submit']").first.click()
    page.wait_for_timeout(3000)

    cancel = page.locator("button:has-text('Cancel')")
    if cancel.count() > 0 and cancel.first.is_visible():
        cancel.first.click()
        page.wait_for_timeout(1500)

    print("  [LOGIN] 完成", flush=True)


# 覆盖项目级 _expose_page_for_report：parameter_settings 用 app_page，不触发 pytest-playwright page
@pytest.fixture(autouse=True)
def _expose_page_for_report(request, app_page: Page):
    request.node.funcargs_page = app_page
    yield


# ── 失败截图（autouse，每条用例自动执行） ─────────────────────────────────────
@pytest.fixture(autouse=True)
def screenshot_on_fail(request, app_page: Page):
    yield
    rep = getattr(request.node, "rep_call", None)
    if rep and rep.failed:
        safe_name = (
            request.node.name
            .replace("[", "_").replace("]", "")
            .replace("/", "_").replace(":", "_")
        )
        path = str(_LOGS_DIR / f"FAIL_{safe_name}.png")
        try:
            data = app_page.screenshot(full_page=True)
            with open(path, "wb") as f:
                f.write(data)
            print(f"\n  [SCREENSHOT] 已保存 → {path}", flush=True)
            allure.attach(data, name="失败截图",
                          attachment_type=allure.attachment_type.PNG)
        except Exception as exc:
            print(f"\n  [SCREENSHOT] 截图失败: {exc}", flush=True)
