# -*- coding: utf-8 -*-
"""_resolve_tcp_connection 纯函数单测：不依赖 Playwright / 网关 / Modbus。

下方 import 依赖本目录 conftest.py 在 collection 前完成的 sys.path 注入
（data_log 目录与仓库根入 path）；从仓库根用 pytest 运行即可。
"""
import config
from datalog_server_verifier import _DISCOVERED_SENTINEL, _resolve_tcp_connection
from projects.AcuHMI_1_7.helpers.physical_devices_reader import DiscoveredDevice


def _dev(name, template, online, ip, unit):
    return DiscoveredDevice(
        name=name, model=template, template=template, online=online,
        serial_id=f"SN-{name}", transport="tcp", ip=ip, port=502, unit=unit,
    )


def test_returns_first_online_device_of_type(monkeypatch):
    monkeypatch.setattr(config, "DISCOVERED_DEVICES", [
        _dev("AcuRev4100_off", "AcuRev4100", False, "10.0.0.1", 1),
        _dev("AcuRev4100_a",   "AcuRev4100", True,  "10.0.0.2", 2),
        _dev("AcuRev4100_b",   "AcuRev4100", True,  "10.0.0.3", 3),
    ], raising=False)
    assert _resolve_tcp_connection("AcuRev4100") == ("10.0.0.2", 502, 2)


def test_returns_none_when_no_discovered_devices(monkeypatch):
    monkeypatch.setattr(config, "DISCOVERED_DEVICES", [], raising=False)
    assert _resolve_tcp_connection("AcuRev4100") is None


def test_returns_none_when_only_offline_or_other_type(monkeypatch):
    monkeypatch.setattr(config, "DISCOVERED_DEVICES", [
        _dev("AcuRev4100_off", "AcuRev4100", False, "10.0.0.1", 1),
        _dev("AcuvimIIW_on",   "AcuvimIIW",  True,  "10.0.0.4", 4),
    ], raising=False)
    assert _resolve_tcp_connection("AcuRev4100") is None


def test_sentinel_is_never_a_static_map_key():
    # 哨兵键绝不能是 MODBUS_DEVICE_MAP 的键，否则 comparator 会误命中静态表
    assert _DISCOVERED_SENTINEL not in config.MODBUS_DEVICE_MAP
