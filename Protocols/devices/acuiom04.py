# -*- coding: utf-8 -*-
"""
devices/acuiom04.py — AcuIOM-04 Modbus 参数地址映射

型号规格：28 DI + 4 DO + 2 RO
地址来源：AcuIOM Modbus Address Table v1.01

DI Status（FC 02H，bit）：0x0000–0x001B（DI 1–28）
DO Status（FC 01H，bit）：0x0000–0x0003（DO 1–4）
RO Status（FC 01H，bit）：0x0020–0x0021（RO 1–2）
DI Pulse Count（FC 03H，uint32）：0x2200–0x2237（DI 1–28，每项 2 寄存器）

BACnet 比对参数范围：模板 blockParams 页 range 列包含 "10" 的参数。
比对参数集合：DI_*_Status（BI 对象）、Conf_DI*_pulse_count（AI 对象）、
              DO*_Status（BO 对象，网关当前不发布 BO，范围检查时显示为"缺失"）、
              RO*_Status（同上）。

对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """AcuIOM-04 param_key → ModbusRegister 映射。"""
    regs: dict[str, ModbusRegister] = {}

    # ── DI Status（FC 02H，bit，addr 0x0000–0x001B） ─────────────────────────
    for i in range(1, 29):
        key = f"DI_{i}_Status"
        regs[key] = ModbusRegister(key, i - 1, 'bit', fc=2)

    # ── DO Status（FC 01H，bit，addr 0x0000–0x0003） ──────────────────────────
    for i in range(1, 5):
        key = f"DO{i}_Status"
        regs[key] = ModbusRegister(key, i - 1, 'bit', fc=1)

    # ── RO Status（FC 01H，bit，addr 0x0020–0x0021） ──────────────────────────
    for i in range(1, 3):
        key = f"RO{i}_Status"
        regs[key] = ModbusRegister(key, 0x0020 + i - 1, 'bit', fc=1)

    # ── DI Pulse Count（FC 03H，uint32，addr 0x2200–0x2237，步长 2） ──────────
    for i in range(1, 29):
        key = f"Conf_DI{i}_pulse_count"
        regs[key] = ModbusRegister(key, 0x2200 + (i - 1) * 2, 'uint32', fc=3)

    return regs
