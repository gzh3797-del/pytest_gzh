# -*- coding: utf-8 -*-
"""
devices/pxm350.py — PXM350 Modbus 参数地址映射

地址来源：AcuRev1310_PXM350_ Modbus Address_v1.01_Sam Xu_260305.xlsx
FC 03H Read Holding Registers，float32 / uint32（int32 按无符号解码）。

BACnet 对象名前缀：PXM350s15-
对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """生成 PXM350 param_key → ModbusRegister 映射。"""
    F32 = 'float32'
    U32 = 'uint32'
    regs: dict[str, ModbusRegister] = {}

    def add(key: str, addr: int, dtype: str = F32) -> None:
        regs[key] = ModbusRegister(key, addr, dtype)

    # ── 实时量（0x2000, float32） ──────────────────────────────────────────────
    # 0x2000 = Total Current（不在 EPICS，跳过）
    add('I_a_A',     0x2002)
    add('I_b_A',     0x2004)
    add('I_c_A',     0x2006)
    add('VLN_avg_V', 0x2008)
    add('VLN_a_V',   0x200A)
    add('VLN_b_V',   0x200C)
    add('VLN_c_V',   0x200E)
    add('VLL_avg_V', 0x2010)
    add('VLL_ab_V',  0x2012)
    add('VLL_bc_V',  0x2014)
    add('VLL_ca_V',  0x2016)
    add('FREQ_Hz',   0x2018)
    add('P_kW',      0x201A)
    add('P_a_kW',    0x201C)
    add('P_b_kW',    0x201E)
    add('P_c_kW',    0x2020)
    add('S_kVA',     0x2022)
    add('S_a_kVA',   0x2024)
    add('S_b_kVA',   0x2026)
    add('S_c_kVA',   0x2028)
    add('Q_kvar',    0x202A)
    add('Q_a_kvar',  0x202C)
    add('Q_b_kvar',  0x202E)
    add('Q_c_kvar',  0x2030)
    add('PF',        0x2032)
    add('PF_a',      0x2034)
    add('PF_b',      0x2036)
    add('PF_c',      0x2038)
    # 0x203A-0x2045 = Phase Angles (not in PXM350s15 EPICS)

    # ── 有功电能（int32，正值等价 uint32，0x2046） ─────────────────────────────
    add('EP_EXP_kWh',   0x2046, U32)   # Total Active Energy Exported
    add('EP_EXP_a_kWh', 0x2048, U32)   # Phase A Export
    add('EP_EXP_b_kWh', 0x204A, U32)   # Phase B Export
    add('EP_EXP_c_kWh', 0x204C, U32)   # Phase C Export
    add('EP_IMP_kWh',   0x204E, U32)   # Total Active Energy Imported
    add('EP_IMP_a_kWh', 0x2050, U32)   # Phase A Import
    add('EP_IMP_b_kWh', 0x2052, U32)   # Phase B Import
    add('EP_IMP_c_kWh', 0x2054, U32)   # Phase C Import

    # ── 需量（地址需从 Excel 精确确认，暂不映射） ─────────────────────────────
    # DMD_P_kW, DMD_Q_kvar, DMD_S_kVA 待补充

    return regs
