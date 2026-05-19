# -*- coding: utf-8 -*-
"""
devices/acuiom01.py — AcuIOM-01 Modbus 参数地址映射

型号规格：8 AI + 2 AO
地址来源：AcuIOM Modbus Address Table v1.01

AI 原始输入数据（FC 03H，float32）：0x3500–0x350F（AI1-8，步长 2）
AI 物理量读数（FC 03H，float32）：0x3700–0x370F（AI1-8，步长 2）
AO 输出数据在 BACnet 中为 AO 对象，当前框架仅支持 AI 对象，故不纳入。

对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """AcuIOM-01 param_key → ModbusRegister 映射（8 个 AI 通道，原始输入 + 物理量读数）。"""
    regs: dict[str, ModbusRegister] = {}

    # ── AI 原始输入数据（0x3500–0x350F，float32） ────────────────────────────
    for i in range(1, 9):
        key = f'AI{i}_input_original_data'
        regs[key] = ModbusRegister(key, 0x3500 + (i - 1) * 2, 'float32')

    # ── AI 物理量读数（0x3700–0x370F，float32，engineering unit） ─────────────
    for i in range(1, 9):
        key = f'AI{i}_physical_measurement_reading'
        regs[key] = ModbusRegister(key, 0x3700 + (i - 1) * 2, 'float32')

    return regs
