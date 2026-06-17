# -*- coding: utf-8 -*-
"""
devices/acuiom02.py — AcuIOM-02 Modbus 参数地址映射

型号规格：16 AI + 4 AO
地址来源：AcuIOM Modbus Address Table v1.01

AI 原始输入数据（FC 03H，float32）：0x3500–0x351F（AI1-16，步长 2）
AI 物理量读数（FC 03H，float32）：0x3700–0x371F（AI1-16，步长 2）
AO 输出数据（FC 03H，float32）：0x3900–0x3907（BACnet 侧为 AO 对象，当前框架不支持）

对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """AcuIOM-02 param_key → ModbusRegister 映射（16 个 AI 通道，原始输入 + 物理量读数）。"""
    regs: dict[str, ModbusRegister] = {}

    # ── AI 原始输入数据（0x3500–0x351F，float32） ────────────────────────────
    for i in range(1, 17):
        key = f'AI{i}_input_original_data'
        regs[key] = ModbusRegister(key, 0x3500 + (i - 1) * 2, 'float32')

    # ── AI 物理量读数（0x3700–0x371F，float32，engineering unit） ─────────────
    for i in range(1, 17):
        key = f'AI{i}_physical_measurement_reading'
        regs[key] = ModbusRegister(key, 0x3700 + (i - 1) * 2, 'float32')

    return regs
