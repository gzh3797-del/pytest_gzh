# -*- coding: utf-8 -*-
"""
conftest.py — acuiom01 子模块专用 fixtures

动态发现网关下挂的在线 AcuIOM-1 设备（Modbus TCP），把 name/ip/port/unit 写入
helpers_iom01 / _src_io_ai / _src_io_ao 的连接常量。与 acurev4100/conftest.py
同一套机制（复用 physical_devices_reader）。
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

_TEMPLATE = "AcuIOM01"


@pytest.fixture(scope="package")
def discovered_devices(app_page: Page) -> list[DiscoveredDevice]:
    """复用已登录的 app_page（父级 conftest 提供），调网关 API 发现下挂 Modbus 设备。"""
    return discover_modbus_tcp_devices(app_page)


@pytest.fixture(scope="package", autouse=True)
def _bind_acuiom01_device(
    discovered_devices: list[DiscoveredDevice],
) -> DiscoveredDevice:
    """整个 acuiom01 会话只解析一次：按模板名取当前在线的 AcuIOM-1 设备，
    把 name/ip/port/unit 写入 helpers_iom01（page 操作模块只依赖 DEVICE_NAME，
    Modbus 校验只依赖 MODBUS_HOST/PORT/SLAVE_ID，均在调用时读取模块全局值）。
    """
    dev = pick_device_for_template(discovered_devices, _TEMPLATE, online_only=True)
    if dev is None:
        pytest.fail(
            f"网关下挂的在线 Modbus TCP 设备中无模板为 {_TEMPLATE!r} 的设备"
            f"（已发现：{[(d.name, d.model, d.online) for d in discovered_devices]}）。"
            "请确认 AcuIOM-1 已上电并接入网关。"
        )
    if dev.transport != "tcp":
        pytest.fail(
            f"发现的 {_TEMPLATE!r} 设备 {dev.name!r} 非 Modbus TCP 接入，"
            "无法直连做寄存器校验（本模块用例依赖直连校验）。"
        )

    import helpers_iom01
    import _src_io_ai  # noqa: F401  仅为触发模块加载，函数按需从 helpers_iom01 取值
    import _src_io_ao  # noqa: F401

    helpers_iom01.DEVICE_NAME = dev.name
    helpers_iom01.MODBUS_HOST = dev.ip
    helpers_iom01.MODBUS_PORT = dev.port
    helpers_iom01.SLAVE_ID = dev.unit
    helpers_iom01._modbus_client = None  # 清掉旧连接缓存，强制按新地址重连

    return dev
