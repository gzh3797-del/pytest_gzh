# -*- coding: utf-8 -*-
"""
conftest.py — pytest fixtures，供 AcuHMI-1-7 BACnet/IP UI 测试使用。

提供：
    hmi_page  — 已登录并导航到 BACnet/IP 配置页面的 Playwright Page 对象

报告：统一使用 pytest-html（运行时由 run.py 或命令行 --html 指定输出，与其他模块一致）。
    - _record_case_id fixture：用例编号写入 JUnit XML property（兼容 CI/CD）。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

import pytest
from playwright.sync_api import Browser, BrowserContext, Page, sync_playwright

# 把仓库根目录加入 path
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 设备 Web UI 配置（统一维护于 projects/AcuHMI_1_7/settings.py）────────────
from projects.AcuHMI_1_7.settings import get_screenshot_dir  # noqa: E402
from projects.AcuHMI_1_7.pages.login_page import LoginPage  # noqa: E402


# ── 进程级 Playwright 实例 ────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def playwright_instance():
    """Session 级 Playwright 实例（避免重复启动浏览器）。"""
    with sync_playwright() as pw:
        yield pw


@pytest.fixture(scope="session")
def browser(playwright_instance) -> Browser:
    """Session 级 Chromium 浏览器（无头模式）。"""
    br = playwright_instance.chromium.launch(headless=False)
    yield br
    br.close()


@pytest.fixture(scope="session")
def browser_context(browser) -> BrowserContext:
    """Session 级浏览器上下文（忽略 HTTPS 证书错误）。"""
    ctx = browser.new_context(ignore_https_errors=True)
    yield ctx
    ctx.close()


# ── 登录并导航到 BACnet/IP 页面 ────────────────────────────────────────────────

def _login_and_navigate_bacnet(page: Page) -> None:
    """登录 HMI Web UI 并导航到 BACnet/IP 配置页面。

    登录复用项目统一的 LoginPage（账号/密码取自 settings 的 DEFAULT_USERNAME/
    DEFAULT_PASSWORD，即 .env 的 WEB_USERNAME/WEB_PASSWORD），与 tests/ui 当前结构
    保持一致；导航选择器同样对齐 tests/ui/protocols 下的 `_nav_protocol` 写法。
    """
    # ── 登录（选择器/凭据与当前结构 LoginPage 一致）──
    login = LoginPage(page)
    login.open()
    login.login()

    # ── 切换到 AcuHMI 设备管理菜单 ──
    hmi_nav = page.locator("header span").filter(has_text="AcuHMI")
    if hmi_nav.count() > 0:
        hmi_nav.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    # ── 左侧导航 Protocols ──
    page.locator(".left-nav-item").filter(has_text="Protocols").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # ── BACnet/IP 子菜单 ──
    page.get_by_role("menuitem", name="BACnet/IP").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


@pytest.fixture(scope="module")
def hmi_page(browser_context) -> Page:
    """Module 级 Page：已登录并在 BACnet/IP 配置页面。"""
    page = browser_context.new_page()
    _login_and_navigate_bacnet(page)
    yield page
    page.close()


# ── pytest-html 报告：用例编号列 ──────────────────────────────────────────────

_CASE_ID_RE = re.compile(r"(TestCase_AcuHMI[\w\-]+)")


@pytest.fixture(autouse=True)
def _record_case_id(request: pytest.FixtureRequest, record_property: Any) -> None:
    """将用例编号写入 JUnit XML property（兼容 CI/CD 系统）。"""
    doc = (getattr(request.node.function, "__doc__", "") or "").strip()
    m = _CASE_ID_RE.match(doc)
    if m:
        record_property("用例编号", m.group(1))


@pytest.fixture(autouse=True)
def _expose_page_for_report():
    """覆盖项目级同名 autouse fixture `_expose_page_for_report(request, page)`。

    项目级版本依赖 pytest-playwright 插件的 `page` 且为 autouse，会对 bacnet 用例也生效，
    从而在 setup 阶段强制初始化插件的 session `playwright`（sync_playwright）。但 bacnet
    模块本身已用自建 `sync_playwright`（hmi_page），同线程同时存在两套 sync_playwright 会触发
    "Sync API inside the asyncio loop"。本覆盖**不引用插件 page**，置空即断开该依赖；
    失败截图由下方 `screenshot_on_failure` 经 hmi_page 处理。
    """
    yield


@pytest.fixture(autouse=True)
def screenshot_on_failure(request: pytest.FixtureRequest):
    """覆盖项目级同名 fixture。

    项目级 `screenshot_on_failure(request, page)` 依赖 pytest-playwright 插件的 `page`，
    而 bacnet 模块用的是自建 `sync_playwright`（hmi_page）。两者同线程会起两个 sync_playwright
    冲突（"Sync API inside the asyncio loop"）。本覆盖**不引用插件 page**，失败时改用本模块
    的 `hmi_page` 截图，从而只保留一套 sync_playwright。
    """
    yield
    rep = getattr(request.node, "rep_call", None)
    if rep is not None and getattr(rep, "failed", False):
        try:
            page = request.getfixturevalue("hmi_page")
            safe = re.sub(r"[^\w.\-]", "_", request.node.name)
            page.screenshot(path=str(get_screenshot_dir() / f"FAIL_{safe}.png"))
        except Exception:  # pragma: no cover - 截图失败不应影响测试结果
            pass
