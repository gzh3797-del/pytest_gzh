"""conftest.py — PX-EMD-G 项目级 pytest fixtures。

迁移自 AcuHMI-1-7，按 BACnet/IP 模块所需精简：仅保留全项目唯一的 Playwright
实例、浏览器 context 参数、报告失败截图钩子。各子模块（如 tests/BacnetIP）在
自己的 conftest 中按需覆盖 fixture。
"""
import base64
import sys
from pathlib import Path

import pytest
from playwright.sync_api import sync_playwright, Browser

from projects.PX_EMD_G.settings import HEADLESS, SLOW_MO, BASE_URL

# 强制标准输出/错误流使用 UTF-8，避免 Windows 控制台中文乱码
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


# ── 全项目唯一 Playwright 实例 ─────────────────────────────────────────────
# 覆盖 pytest-playwright ≥0.5 内置的 `playwright` / `browser` fixture，让所有子模块
# 共享同一实例，避免多处 sync_playwright() 在同一线程内竞争事件循环。

@pytest.fixture(scope="session")
def playwright():
    with sync_playwright() as pw:
        yield pw


@pytest.fixture(scope="session")
def browser(playwright) -> Browser:
    br = playwright.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
    yield br
    br.close()


@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    return {
        **browser_context_args,
        "base_url": BASE_URL,
        "ignore_https_errors": True,
        "viewport": {"width": 1280, "height": 720},
    }


# ── Page Object fixture ─────────────────────────────────────────────────────

@pytest.fixture
def login_page(page):
    from projects.PX_EMD_G.pages.login_page import LoginPage
    return LoginPage(page)


# ── 把 page 暴露给报告钩子（供失败时截图）──────────────────────────────────

@pytest.fixture(autouse=True)
def _expose_page_for_report(request):
    # autouse：让 pytest_runtest_makereport 在用例失败时能取到 page 截图。
    # 仅当用例已通过其他 fixture 持有 page 时才暴露，不主动创建新 context。
    # 各子模块可在自己的 conftest 覆盖此 fixture，提供对应的 page 对象。
    if "app_page" in request.fixturenames:
        request.node.funcargs_page = request.getfixturevalue("app_page")
    yield


# ── 报告增强：失败错误信息（traceback 原生 + 截图内嵌）──────────────────────

@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    # 每个阶段生成报告对象后挂到 item（item.rep_call 等），供子模块的 screenshot_on_failure
    # 读取；call 阶段失败时把截图 base64 内嵌进 pytest-html 报告对应行。
    pytest_html = item.config.pluginmanager.getplugin("html")
    outcome = yield
    rep = outcome.get_result()
    setattr(item, f"rep_{rep.when}", rep)

    if pytest_html is None:
        return

    extras = list(getattr(rep, "extras", []))

    if rep.when == "call" and rep.failed:
        page = getattr(item, "funcargs_page", None)
        if page is not None:
            try:
                from projects.PX_EMD_G.settings import get_screenshot_dir
                from projects.PX_EMD_G.helpers.web_helpers import timestamp
                png = page.screenshot()
                shot_path = get_screenshot_dir() / f"FAIL_{item.name}_{timestamp()}.png"
                shot_path.write_bytes(png)
                b64 = base64.b64encode(png).decode("ascii")
                extras.append(pytest_html.extras.image(b64, name="失败截图", mime_type="image/png"))
            except Exception:
                pass  # 截图失败不应影响报告生成

    rep.extras = extras
