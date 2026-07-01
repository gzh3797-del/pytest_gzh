# -*- coding: utf-8 -*-
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../../.."))

from projects.AcuHMI_1_7.helpers.physical_devices_reader import (
    DiscoveredDevice,
    build_device,
    connection_map,
    pick_device_for_template,
)

# 真实 list 接口样例（deviceType==3 为 Modbus TCP）
_LIST_4100 = {
    "deviceName": "Acurev4100242", "serialNumber": "MODTCP-aaa",
    "isOnline": True, "deviceType": 3, "deviceModel": "AcuRev-4110-mA",
}
# 真实 config 接口样例
_CFG_4100 = {
    "modbusTCPConfig": {"ipAddr": "192.168.2.242", "port": 502, "slaveAddress": 23},
}


def test_build_device_parses_tcp_entry():
    d = build_device(_LIST_4100, _CFG_4100)
    assert d == DiscoveredDevice(
        name="Acurev4100242", model="AcuRev-4110-mA", template="AcuRev4100",
        online=True, serial_id="MODTCP-aaa", transport="tcp",
        ip="192.168.2.242", port=502, unit=23,
    )


def test_build_device_unknown_model_keeps_none_template():
    entry = dict(_LIST_4100, deviceModel="Mystery")
    d = build_device(entry, _CFG_4100)
    assert d is not None and d.template is None


def test_build_device_missing_modbus_cfg_returns_none():
    assert build_device(_LIST_4100, {"modbusTCPConfig": {}}) is None
    assert build_device(_LIST_4100, {}) is None


def test_connection_map_shape():
    d = build_device(_LIST_4100, _CFG_4100)
    assert connection_map([d]) == {"Acurev4100242": ("192.168.2.242", 502, 23)}


def test_pick_first_online_of_template():
    on1 = DiscoveredDevice("A", "AcuRev-4110-mA", "AcuRev4100", True, "s1", "tcp", "1.1.1.1", 502, 1)
    on2 = DiscoveredDevice("B", "AcuRev-4110-mA", "AcuRev4100", True, "s2", "tcp", "1.1.1.2", 502, 2)
    off = DiscoveredDevice("C", "AcuRev-4110-mA", "AcuRev4100", False, "s3", "tcp", "1.1.1.3", 502, 3)
    assert pick_device_for_template([off, on1, on2], "AcuRev4100").name == "A"
    assert pick_device_for_template([off], "AcuRev4100") is None  # 全离线
    assert pick_device_for_template([on1], "AcuvimIIW") is None    # 无匹配模板
