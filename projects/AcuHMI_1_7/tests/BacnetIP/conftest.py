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
from playwright.sync_api import Browser, Page

# 把仓库根目录加入 path
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# ── 设备 Web UI 配置（统一维护于 projects/AcuHMI_1_7/settings.py）────────────
from projects.AcuHMI_1_7.settings import get_screenshot_dir  # noqa: E402
from projects.AcuHMI_1_7.pages.login_page import LoginPage  # noqa: E402
from projects.AcuHMI_1_7.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
    discover_modbus_tcp_devices,
)


# ── 登录并导航到 BACnet/IP 页面 ────────────────────────────────────────────────

def _login_and_navigate_bacnet(page: Page) -> None:
    """登录 HMI Web UI 并导航到 BACnet/IP 配置页面。

    登录复用项目统一的 LoginPage（账号/密码取自 settings 的 DEFAULT_USERNAME/
    DEFAULT_PASSWORD，即 .env 的 WEB_USERNAME/WEB_PASSWORD），与 tests/ui 当前结构
    保持一致。

    导航使用 JS evaluate() click（与 _navigate_to_bacnet 保持一致）：
    直接 Playwright click 在 SPA 初始化期间因 overlay/未渲染等原因容易 30s 超时，
    JS click 不等 actionability，与 _navigate_to_bacnet 风格统一。
    """
    # ── 登录（选择器/凭据与当前结构 LoginPage 一致）──
    login = LoginPage(page)
    login.open()
    login.login()

    # ── 等待登录后 SPA 主框架渲染就绪（最长 15s），避免下面的导航点击落空 ──
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(800)

    # ── 切换到 AcuHMI 设备管理菜单（JS click，绕过 overlay / actionability 检查）──
    page.evaluate(
        """() => {
            for (const el of document.querySelectorAll('header span, .nav-item')) {
                if (el.textContent.includes('AcuHMI')) { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(1200)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass

    # ── 左侧导航 Protocols（JS click）──
    page.evaluate(
        """() => {
            for (const el of document.querySelectorAll('.left-nav-item')) {
                if (el.textContent.trim() === 'Protocols') { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(1000)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass

    # ── BACnet/IP 子菜单（JS click，兼容 menuitem / li 两种结构）──
    page.evaluate(
        """() => {
            for (const el of document.querySelectorAll('[role="menuitem"], li')) {
                if (el.textContent.trim() === 'BACnet/IP') { el.click(); return; }
            }
        }"""
    )
    page.wait_for_timeout(2000)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass


@pytest.fixture(scope="module")
def hmi_page(browser: Browser) -> Page:
    """Module 级 Page：已登录并在 BACnet/IP 配置页面。"""
    ctx = browser.new_context(ignore_https_errors=True)
    page = ctx.new_page()
    _login_and_navigate_bacnet(page)
    yield page
    ctx.close()


@pytest.fixture(scope="module")
def discovered_devices(hmi_page: Page) -> list[DiscoveredDevice]:
    """复用已登录的 hmi_page（sessionStorage 含 token），调网关 API 发现下挂 Modbus TCP 设备。"""
    return discover_modbus_tcp_devices(hmi_page)


# ── pytest-html 报告：用例编号列 ──────────────────────────────────────────────

_CASE_ID_RE = re.compile(r"(TestCase_AcuHMI[\w\-]+)")


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    """把用例编号注入 BacnetIP 用例的 nodeid，让 pytest-html 报告的 Test 列直接显示编号。

    BacnetIP 用例不是参数化用例，nodeid 默认只含 `文件::类::函数名`（如
    `test_019_acurev2100_param_list_matches_template`），报告里看不到标准用例编号。
    用例编号写在每个用例 docstring 首段（如 `TestCase_AcuHMI-1-7_033_001_019: ...`），
    在收集阶段读出后以 `[编号]` 形式追加到 nodeid 末尾，报告即显示
    `test_019_..._matches_template[TestCase_AcuHMI-1-7_033_001_019]`，
    与接线检查模块（parametrize ids 前缀编号）的呈现方式一致。

    子目录 conftest 的 pytest_collection_modifyitems 会收到**全量** items，故先按本
    conftest 所在目录过滤，只改 BacnetIP 自己的用例；并对结尾做幂等判断，避免重复追加。
    """
    here = Path(__file__).parent.resolve()
    for item in items:
        try:
            item_path = Path(str(item.fspath)).resolve()
        except Exception:
            continue
        if here != item_path.parent and here not in item_path.parents:
            continue
        doc = (getattr(getattr(item, "function", None), "__doc__", "") or "").strip()
        m = _CASE_ID_RE.match(doc)
        if not m:
            continue
        case_id = m.group(1)
        suffix = f"[{case_id}]"
        if not item._nodeid.endswith(suffix):
            item._nodeid = f"{item._nodeid}{suffix}"


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

    项目级版本对无 app_page 的用例会请求 pytest-playwright 的 `page` fixture，触发额外的
    browser context 创建。BacnetIP 用例通过 hmi_page 管理自己的 context/page，不需要额外
    的 `page` fixture。本覆盖置空即断开该依赖；失败截图由下方 `screenshot_on_failure` 处理。
    """
    yield


@pytest.fixture(autouse=True)
def screenshot_on_failure(request: pytest.FixtureRequest):
    """覆盖项目级同名 fixture。

    项目级版本对无 app_page 的用例会请求 pytest-playwright 的 `page` fixture，触发额外的
    context 创建。BacnetIP 用例通过 hmi_page 管理自己的 page，失败截图直接从 hmi_page 取，
    不需要额外的 `page` fixture。
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
