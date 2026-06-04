# -*- coding: utf-8 -*-
"""
conftest.py — pytest fixtures，供 AcuHMI-1-7 BACnet/IP UI 测试使用。

提供：
    hmi_page  — 已登录并导航到 BACnet/IP 配置页面的 Playwright Page 对象

报告集成：
    pytest_runtest_makereport / pytest_html_results_table_* 钩子
    从每个测试函数的 docstring 首行提取 TestCase_AcuHMI-1-7_* 编号，
    注入 pytest-html 报告的「用例编号」列和 JUnit XML property。
"""
from __future__ import annotations

import re
import sys
import time
from pathlib import Path
from typing import Any

import pytest
from playwright.sync_api import Browser, BrowserContext, Page, sync_playwright

# 把仓库根目录加入 path
_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 设备 Web UI 配置（统一维护于 test_case/AcuHMI_1_7/config.py）────────────
from test_case.AcuHMI_1_7.config import HMI_URL, HMI_USERNAME, HMI_PASSWORD  # noqa: E402

SCREENSHOTS_DIR = Path(__file__).parent / "screenshots"
SCREENSHOTS_DIR.mkdir(exist_ok=True)


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
    """登录 HMI Web UI 并导航到 BACnet/IP 配置页面。"""
    page.goto(f"{HMI_URL}/#/login", timeout=30000)
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    time.sleep(1)

    page.fill('input[type="text"]', HMI_USERNAME)
    page.fill('input[type="password"]', HMI_PASSWORD)
    page.click('button:has-text("Sign in")')

    try:
        page.wait_for_url(lambda u: "login" not in u, timeout=15000)
    except Exception:
        pass
    page.wait_for_load_state("domcontentloaded", timeout=10000)
    time.sleep(2)

    # Click "AcuHMI-1-7" top nav to switch to device management menu
    hmi_nav = page.locator('.nav-item:has-text("AcuHMI-1-7")')
    if hmi_nav.count() > 0:
        hmi_nav.first.click()
        time.sleep(1)

    # Click "Protocols" left nav
    protocols = page.locator('.left-nav-item:has-text("Protocols")')
    if protocols.count() > 0:
        protocols.first.click()
        time.sleep(1)
        page.wait_for_load_state("domcontentloaded", timeout=8000)

    # Click "BACnet/IP" sub-menu
    bacnet = page.locator('li:has-text("BACnet/IP")')
    if bacnet.count() > 0:
        bacnet.first.click()
        time.sleep(2)
        page.wait_for_load_state("domcontentloaded", timeout=8000)


@pytest.fixture(scope="module")
def hmi_page(browser_context) -> Page:
    """Module 级 Page：已登录并在 BACnet/IP 配置页面。"""
    page = browser_context.new_page()
    _login_and_navigate_bacnet(page)
    yield page
    page.close()


# ── pytest-html 报告：用例编号列 ──────────────────────────────────────────────

_CASE_ID_RE = re.compile(r"(TestCase_AcuHMI[\w\-]+)")


def _extract_case_id(item: Any) -> str:
    """从测试函数 docstring 首行提取 TestCase_AcuHMI-* 格式的用例编号。"""
    fn = getattr(item, "function", None)
    doc = (getattr(fn, "__doc__", "") or "").strip()
    m = _CASE_ID_RE.match(doc)
    return m.group(1) if m else ""


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item: Any, call: Any) -> Any:
    """将用例编号附加到 report 对象，供 pytest-html 行钩子读取。"""
    outcome = yield
    report = outcome.get_result()
    report.case_id = _extract_case_id(item)  # type: ignore[attr-defined]


def pytest_html_results_table_header(cells: list) -> None:
    """在 pytest-html 报告表头第 3 列插入「用例编号」。"""
    try:
        from py.xml import html  # pytest-html v3
        cells.insert(2, html.th("用例编号", class_="sortable"))
    except ImportError:
        cells.insert(2, "<th>用例编号</th>")


def pytest_html_results_table_row(report: Any, cells: list) -> None:
    """在 pytest-html 报告每行第 3 列插入用例编号。"""
    case_id = getattr(report, "case_id", "") or "—"
    try:
        from py.xml import html  # pytest-html v3
        cells.insert(2, html.td(case_id))
    except ImportError:
        cells.insert(2, f"<td>{case_id}</td>")


@pytest.fixture(autouse=True)
def _record_case_id(request: pytest.FixtureRequest, record_property: Any) -> None:
    """将用例编号写入 JUnit XML property（兼容 CI/CD 系统）。"""
    doc = (getattr(request.node.function, "__doc__", "") or "").strip()
    m = _CASE_ID_RE.match(doc)
    if m:
        record_property("用例编号", m.group(1))
