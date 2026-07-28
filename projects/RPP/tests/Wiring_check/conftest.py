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

from projects.RPP.tests.Wiring_check.core import config as cfg
from projects.RPP.tests.Wiring_check.core import report as rpt
from projects.RPP.tests.Wiring_check.core.meter_modbus import WiringCheckModbus
from projects.RPP.tests.Wiring_check.core.wiring_check_page import WiringCheckPage
from projects.RPP.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
    discover_modbus_tcp_devices,
)

# ── 覆盖项目级 autouse fixture，断开对插件 page 的依赖 ─────────────────────────
# 项目级 `_expose_page_for_report(request, page)` 是 autouse，依赖 pytest-playwright
# 插件的 page fixture，会对接线检查用例也生效，从而在接线检查用例的 setup 阶段强制
# 拉起插件内部的 page 创建链。置空覆盖即断开该依赖，避免不必要的 page 初始化。

@pytest.fixture(autouse=True)
def _expose_page_for_report():
    yield


# ── 浏览器（session 级）──────────────────────────────────────────────────────

@pytest.fixture(scope='session')
def _playwright(playwright):
    # 复用项目级唯一 Playwright 实例，避免同线程第二个 sync_playwright 竞争事件循环
    yield playwright


@pytest.fixture(scope='session')
def wc_browser(_playwright):
    browser = _playwright.chromium.launch(headless=False)
    yield browser
    browser.close()


@pytest.fixture(scope='session')
def wc_page(wc_browser):
    """登录并导航到接线检查页面，整个会话共用"""
    from projects.RPP.tests.Wiring_check.test_3e4wy import DEVICE_NAME
    page = wc_browser.new_page(ignore_https_errors=True)
    wcp = WiringCheckPage(page, device_name=DEVICE_NAME)
    wcp.login_if_needed()
    wcp.navigate()
    yield wcp


# ── Modbus 连接（session 级）─────────────────────────────────────────────────

@pytest.fixture(scope='session', autouse=True)
def meter_connection(wc_page) -> DiscoveredDevice:
    """从网关动态发现被测表（cfg.WIRING_DEVICE_NAME）连接参数并填入 cfg。

    复用已登录的 wc_page，按设备名精确匹配；找不到 / 离线时立即抛清晰错误。
    """
    devices = discover_modbus_tcp_devices(wc_page.page)
    target = next((d for d in devices if d.name == cfg.WIRING_DEVICE_NAME), None)
    if target is None:
        names = [d.name for d in devices if d.online]
        raise RuntimeError(
            f"meter_device_name={cfg.WIRING_DEVICE_NAME!r} 未在网关下挂 Modbus TCP 设备中找到。"
            f"当前在线设备：{names}。请把 config.yaml 的 meter_device_name 改为其中之一。"
        )
    if not target.online:
        raise RuntimeError(
            f"被测表 {cfg.WIRING_DEVICE_NAME!r} 当前离线，无法作为接线检查目标。"
        )
    cfg.apply_discovered_connection(target.ip, target.port, target.unit)
    return target


@pytest.fixture(scope='session')
def wiring_modbus(meter_connection):
    """接线检查专用 Modbus TCP 连接，自动写入额定电压"""
    m = WiringCheckModbus()
    m.write_nominal_voltage(cfg.NOMINAL_VOLTAGE)
    yield m
    m.close()


# ── 自定义 HTML 报告收集（pytest 路径）───────────────────────────────────────
# 独立跑 run_all() 会调 report.generate 出 HTML，但 pytest 跑各 test_*.py 不会。
# 这里在 pytest 路径补齐：每条用例把结果记入 session 级收集器，会话结束时按接线方式
# 各生成一份与 run_all 完全相同的自定义 HTML 报告，落在 Wiring_check/reports/。
# 每个 test_*.py 用模块级常量 REPORT_META 声明本接线方式的报告参数（wiring_type 及
# 传给 report.generate 的额外 kwargs，如 active_channels/channel_phases/...）。

@pytest.fixture(scope='session', autouse=True)
def _wc_report_store():
    """汇总各接线方式用例结果，会话结束时生成自定义 HTML 报告。"""
    store: dict[str, dict] = {}
    yield store
    for wiring_type, bundle in store.items():
        results = bundle['results']
        if not results:
            continue
        extra = {k: v for k, v in bundle['meta'].items() if k != 'wiring_type'}
        try:
            path = rpt.generate(
                results=results, wiring_type=wiring_type,
                device_name=cfg.WIRING_DEVICE_NAME, meter_ip=cfg.METER_TCP_IP,
                **extra)
            print(f'\n[wiring-report] {wiring_type} → {path}')
        except Exception as exc:  # 报告生成失败不应影响测试结论
            print(f'\n[wiring-report] {wiring_type} 生成失败：{exc}')


@pytest.fixture
def wc_record(request, _wc_report_store):
    """把单条用例结果记入报告收集器；接线方式与报告参数取自模块级 REPORT_META。"""
    def _record(result: dict):
        meta = getattr(request.module, 'REPORT_META', None)
        if not meta:
            return
        bundle = _wc_report_store.setdefault(
            meta['wiring_type'], {'meta': meta, 'results': []})
        bundle['results'].append(result)

    return _record
