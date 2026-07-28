# -*- coding: utf-8 -*-
"""
conftest.py — acuiom03 子模块专用 fixtures

动态发现网关下挂的 AcuIOM-3 设备（Modbus TCP）。**探查阶段实测该台
（IOM03P170S04）当前离线**——与其余三台 IOM 不同，本 conftest 在设备离线时
调用 pytest.skip() 跳过整组用例（而非 pytest.fail），因为这是已知的现场状态，
不是脚本/环境错误；待设备真正上线后，本组用例会自动开始执行。
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

_TEMPLATE = "AcuIOM03"


@pytest.fixture(scope="package")
def discovered_devices(app_page: Page) -> list[DiscoveredDevice]:
    """复用已登录的 app_page（父级 conftest 提供），调网关 API 发现下挂 Modbus 设备。"""
    return discover_modbus_tcp_devices(app_page)


@pytest.fixture(scope="package", autouse=True)
def _bind_acuiom03_device(
    discovered_devices: list[DiscoveredDevice],
) -> DiscoveredDevice:
    """整个 acuiom03 会话只解析一次：按模板名取 AcuIOM-3 设备。

    - 在线：把 name/ip/port/unit 写入 helpers_iom03，用例正常执行。
    - 发现记录里存在但离线：pytest.skip 整组（已知现场状态）。
    - 发现记录里完全没有该模板：pytest.fail（模板映射/接线异常，需人工排查）。
    """
    dev = pick_device_for_template(discovered_devices, _TEMPLATE, online_only=True)
    if dev is None:
        dev_offline = pick_device_for_template(discovered_devices, _TEMPLATE, online_only=False)
        if dev_offline is not None:
            pytest.skip(
                f"AcuIOM-3 设备 {dev_offline.name!r} 当前离线（探查阶段已知状态，"
                "IOM03P170S04 @ https://192.168.3.47），本组用例需设备上线后执行。"
            )
        pytest.fail(
            f"网关下挂设备中完全找不到模板为 {_TEMPLATE!r} 的 AcuIOM-3"
            f"（已发现：{[(d.name, d.model, d.online) for d in discovered_devices]}）。"
            "请确认 template_matcher.MODEL_TO_TEMPLATE 映射 或 AcuIOM-3 是否已从网关移除。"
        )
    if dev.transport != "tcp":
        pytest.fail(
            f"发现的 {_TEMPLATE!r} 设备 {dev.name!r} 非 Modbus TCP 接入，"
            "无法直连做寄存器校验（本模块用例依赖直连校验）。"
        )

    import helpers_iom03
    import _src_io_di  # noqa: F401
    import _src_io_do  # noqa: F401
    import _src_io_ro  # noqa: F401

    helpers_iom03.DEVICE_NAME = dev.name
    helpers_iom03.MODBUS_HOST = dev.ip
    helpers_iom03.MODBUS_PORT = dev.port
    helpers_iom03.SLAVE_ID = dev.unit
    helpers_iom03._modbus_client = None

    return dev
