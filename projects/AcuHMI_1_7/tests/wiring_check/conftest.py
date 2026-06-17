"""
接线检查测试 pytest fixtures

scope 说明：
  session  — 浏览器实例、Modbus 连接（整个测试会话只建一次）
  module   — 每个 test_*.py 各执行一次 service config 写入
  function — 每条用例独立执行（由 parametrize 驱动）
"""
import sys
import os
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))

from playwright.sync_api import sync_playwright as _sync_playwright
from projects.AcuHMI_1_7.tests.wiring_check.core import config as cfg
from projects.AcuHMI_1_7.tests.wiring_check.core.meter_modbus import WiringCheckModbus
from projects.AcuHMI_1_7.tests.wiring_check.core.wiring_check_page import WiringCheckPage

# ── 覆盖项目级 autouse fixture，断开对插件 page 的依赖 ─────────────────────────
# 项目级 `_expose_page_for_report(request, page)` 是 autouse，依赖 pytest-playwright
# 插件的 page，会对接线检查用例也生效，从而强制初始化插件 session `playwright`
# （sync_playwright）。但本模块用自建 sync_playwright（wc_page），同线程两套 sync_playwright
# 会触发 "Sync API inside the asyncio loop"。置空覆盖即断开该依赖。

@pytest.fixture(autouse=True)
def _expose_page_for_report():
    yield


# ── 浏览器（session 级）──────────────────────────────────────────────────────

@pytest.fixture(scope='session')
def _playwright():
    with _sync_playwright() as pw:
        yield pw


@pytest.fixture(scope='session')
def wc_browser(_playwright):
    browser = _playwright.chromium.launch(headless=False)
    yield browser
    browser.close()


@pytest.fixture(scope='session')
def wc_page(wc_browser):
    """登录并导航到接线检查页面，整个会话共用"""
    from projects.AcuHMI_1_7.tests.wiring_check.test_3e4wy import DEVICE_NAME
    page = wc_browser.new_page(ignore_https_errors=True)
    wcp = WiringCheckPage(page, device_name=DEVICE_NAME)
    wcp.login_if_needed()
    wcp.navigate()
    yield wcp


# ── Modbus 连接（session 级）─────────────────────────────────────────────────

@pytest.fixture(scope='session')
def wiring_modbus():
    """接线检查专用 Modbus TCP 连接，自动写入额定电压"""
    m = WiringCheckModbus()
    m.write_nominal_voltage(cfg.NOMINAL_VOLTAGE)
    yield m
    m.close()
