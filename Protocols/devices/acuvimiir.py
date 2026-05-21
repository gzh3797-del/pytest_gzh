# -*- coding: utf-8 -*-
"""
devices/acuvimiir.py — AcuvimIIR Modbus 参数地址映射

Modbus 地址表与 AcuvimIIW 相同（共用：Acuvim IIW&IIR&CL&EL Modbus Address）。
BACnet 对象名前缀：PXE114-（与 AcuvimIIW- 不同）

对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
# AcuvimIIR 与 AcuvimIIW Modbus 地址完全相同，直接复用
from devices.acuvimiiw import build_param_map, build_cloud_col_map  # noqa: F401
