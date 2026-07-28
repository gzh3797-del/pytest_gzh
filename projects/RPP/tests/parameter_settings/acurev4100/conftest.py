# -*- coding: utf-8 -*-
"""
conftest.py — acurev4100 子模块专用 fixtures

动态发现网关下挂的在线 AcuRev4100 设备（Modbus TCP），替代原来从 config.yaml
device_modbus 段读取的静态 ip/port/unit 配置——避免物理设备切换/离线后，脚本
仍连着一台已经不在线的表（与 tests/BacnetIP 的发现机制一致，复用同一套
physical_devices_reader 工具）。

- discovered_devices        : package 级，复用父级 conftest 的 app_page 发现下挂设备
- _bind_acurev4100_device   : package 级 autouse，整个会话只解析一次，把发现结果
                              写入 helpers_4100 / _src_event_waveform /
                              _src_general_settings 三个模块的连接常量
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page

_REPO_ROOT = Path(__file__).resolve().parents[5]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))
_THIS_DIR = Path(__file__).parent
if str(_THIS_DIR) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR))

from projects.RPP.helpers.physical_devices_reader import (  # noqa: E402
    DiscoveredDevice,
    discover_modbus_tcp_devices,
    pick_device_for_template,
)

_TEMPLATE = "AcuRev4100"


@pytest.fixture(scope="package")
def discovered_devices(app_page: Page) -> list[DiscoveredDevice]:
    """复用已登录的 app_page（父级 conftest 提供），调网关 API 发现下挂 Modbus 设备。"""
    return discover_modbus_tcp_devices(app_page)


@pytest.fixture(scope="package", autouse=True)
def _bind_acurev4100_device(
    discovered_devices: list[DiscoveredDevice],
) -> DiscoveredDevice:
    """整个 acurev4100 会话只解析一次：按模板名取当前在线的 AcuRev4100 设备，
    把 name/ip/port/unit 写入 helpers_4100、_src_event_waveform、
    _src_general_settings 三个模块的连接常量（三者重复定义了同一套常量）。

    35 个 test_*.py 不受影响：它们只调用这三个模块内的函数，不直接读常量，
    函数体内按调用时的模块全局命名空间取值，覆盖后自然生效。
    """
    dev = pick_device_for_template(discovered_devices, _TEMPLATE, online_only=True)
    if dev is None:
        pytest.fail(
            f"网关下挂的在线 Modbus TCP/RTU 设备中无模板为 {_TEMPLATE!r} 的设备"
            f"（已发现：{[(d.name, d.model, d.online) for d in discovered_devices]}）。"
            "请确认该表已上电并接入网关。"
        )
    if dev.transport != "tcp":
        pytest.fail(
            f"发现的 {_TEMPLATE!r} 设备 {dev.name!r} 是 Modbus RTU 接入，"
            "无法直连 Modbus TCP 做寄存器校验（此模块的用例依赖直连校验）。"
        )

    import helpers_4100
    import _src_event_waveform
    import _src_general_settings

    for mod in (helpers_4100, _src_event_waveform, _src_general_settings):
        mod.DEVICE_NAME = dev.name
        mod.MODBUS_HOST = dev.ip
        mod.MODBUS_PORT = dev.port
        mod.SLAVE_ID = dev.unit
        mod._modbus_client = None  # 清掉旧连接缓存，强制按新地址重连

    return dev


@pytest.fixture
def restore_wiring(app_page: Page):
    """接线方式(Wiring)类用例专用：用例结束后把 Wiring 还原为默认 3 Element 4 Wire Y，
    避免把设备遗留在非默认接线态。仅被 test_..._001_07_case38~41 显式请求，不影响其余用例。
    还原失败只告警不使套件失败（teardown 阶段）。"""
    yield
    try:
        import _src_user_and_ct as _uct
        _uct.nav_to_user_and_ct(app_page)
        _uct.switch_wiring(app_page, _uct.DEFAULT_WIRING)
    except Exception as exc:  # noqa: BLE001  teardown 兜底，不因还原失败挂掉用例
        import logging
        logging.getLogger("acurev4100_test").warning("restore_wiring 还原失败: %s", exc)
