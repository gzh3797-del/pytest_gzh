"""
接线检查测试 pytest fixtures

scope 说明：
  session  — 浏览器实例、Modbus 连接（整个测试会话只建一次）
  module   — 每个 test_*.py 各执行一次 service config 写入
  function — 每条用例独立执行（由 parametrize 驱动）

报告：
  pytest-html  → pytest ... --html=reports/report.html --self-contained-html
  自定义HTML   → 每个 test_*.py 直接运行时自动生成（含每通道列）
"""
import sys
import os
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../..'))

from playwright.sync_api import sync_playwright as _sync_playwright
from projects.ACM_41_WEB2.wiring_check.core import config as cfg
from projects.ACM_41_WEB2.wiring_check.core.meter_modbus import WiringCheckModbus
from projects.ACM_41_WEB2.wiring_check.core.wiring_check_page import WiringCheckPage

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
    from projects.ACM_41_WEB2.wiring_check.test_3e4wy import DEVICE_NAME
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
