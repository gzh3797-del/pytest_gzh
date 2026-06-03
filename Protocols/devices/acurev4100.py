# -*- coding: utf-8 -*-
"""
devices/acurev4100.py — AcuRev4100 Modbus 参数地址映射

地址来源：AcuRev4100 Modbus Address Table v1.02 20260202.xlsx
FC 03H Read Holding Registers，float32 高字优先，float64 大端序。

对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """生成 AcuRev4100 完整的 param_key → ModbusRegister 映射（1869 个参数）。"""
    F32 = 'float32'
    F64 = 'float64'
    regs: dict[str, ModbusRegister] = {}

    def add(key: str, addr: int, dtype: str = F32, fc: int = 3) -> None:
        regs[key] = ModbusRegister(key, addr, dtype, fc)

    # ── 系统基础量 (0x2000) ───────────────────────────────────────────────────
    add('FREQ_Hz',   0x2000)
    add('VLN_a_V',   0x2002)
    add('VLN_b_V',   0x2004)
    add('VLN_c_V',   0x2006)
    add('VLN_avg_V', 0x2008)
    add('VLL_ab_V',  0x200A)
    add('VLL_bc_V',  0x200C)
    add('VLL_ca_V',  0x200E)
    add('VLL_avg_V', 0x2010)
    add('ANG_VLN_a', 0x2012)
    add('ANG_VLN_b', 0x2014)
    add('ANG_VLN_c', 0x2016)

    # 各相 I/P/Q/S/PF（相间有 2 个寄存器的 Load Nature 占位）
    for phase, base in [('a', 0x2018), ('b', 0x2024), ('c', 0x2030)]:
        add(f'I_{phase}_A',    base + 0)
        add(f'P_{phase}_kW',   base + 2)
        add(f'Q_{phase}_kvar', base + 4)
        add(f'S_{phase}_kVA',  base + 6)
        add(f'PF_{phase}',     base + 10)  # base+8 = Load Nature，跳过

    # 系统合计
    add('I_avg_A', 0x203C)
    add('P_kW',    0x203E)
    add('Q_kvar',  0x2040)
    add('S_kVA',   0x2042)
    add('PF',      0x2046)  # 0x2044 = System Load Nature，跳过

    # ── 输入通道 001-024 (0x2048, 每通道 14 寄存器) ───────────────────────────
    for ch in range(1, 25):
        n = f'{ch:03d}'
        base = 0x2048 + (ch - 1) * 14
        add(f'I_{n}_A',    base + 0)
        add(f'P_{n}_kW',   base + 2)
        add(f'Q_{n}_kvar', base + 4)
        add(f'S_{n}_kVA',  base + 6)
        add(f'PF_{n}',     base + 10)   # base+8 = Load Nature，跳过
        add(f'ANG_I_{n}',  base + 12)

    # ── 用户通道 s001-s012 (0x2198, 每通道 12 寄存器) ────────────────────────
    for ch in range(1, 13):
        n = f's{ch:03d}'
        base = 0x2198 + (ch - 1) * 12
        add(f'I_{n}_A',    base + 0)
        add(f'P_{n}_kW',   base + 2)
        add(f'Q_{n}_kvar', base + 4)
        add(f'S_{n}_kVA',  base + 6)
        add(f'PF_{n}',     base + 10)   # base+8 = Load Nature，跳过

    # ── 需量 (0x2300) ─────────────────────────────────────────────────────────
    for phase, base in [('a', 0x2300), ('b', 0x230C), ('c', 0x2318)]:
        add(f'DMD_I_{phase}_A',        base + 0)
        add(f'DMD_IMP_P_{phase}_kW',   base + 2)
        add(f'DMD_EXP_P_{phase}_kW',   base + 4)
        add(f'DMD_IMP_Q_{phase}_kvar', base + 6)
        add(f'DMD_EXP_Q_{phase}_kvar', base + 8)
        add(f'DMD_S_{phase}_kVA',      base + 10)

    # 系统需量
    add('DMD_I_avg_A',    0x2324)
    add('DMD_IMP_P_kW',   0x2326)
    add('DMD_EXP_P_kW',   0x2328)
    add('DMD_IMP_Q_kvar', 0x232A)
    add('DMD_EXP_Q_kvar', 0x232C)
    add('DMD_S_kVA',      0x232E)

    # 输入通道需量 (0x2330, 每通道 12 寄存器)
    for ch in range(1, 25):
        n = f'{ch:03d}'
        base = 0x2330 + (ch - 1) * 12
        add(f'DMD_I_{n}_A',        base + 0)
        add(f'DMD_IMP_P_{n}_kW',   base + 2)
        add(f'DMD_EXP_P_{n}_kW',   base + 4)
        add(f'DMD_IMP_Q_{n}_kvar', base + 6)
        add(f'DMD_EXP_Q_{n}_kvar', base + 8)
        add(f'DMD_S_{n}_kVA',      base + 10)

    # 用户通道需量 (0x2450, 每通道 12 寄存器)
    for ch in range(1, 13):
        n = f's{ch:03d}'
        base = 0x2450 + (ch - 1) * 12
        add(f'DMD_I_{n}_A',        base + 0)
        add(f'DMD_IMP_P_{n}_kW',   base + 2)
        add(f'DMD_EXP_P_{n}_kW',   base + 4)
        add(f'DMD_IMP_Q_{n}_kvar', base + 6)
        add(f'DMD_EXP_Q_{n}_kvar', base + 8)
        add(f'DMD_S_{n}_kVA',      base + 10)

    # ── 电能质量 (0x2600) ─────────────────────────────────────────────────────
    add('MAG_SEQ_POS_V',  0x2600)
    add('MAG_SEQ_ZERO_V', 0x2602)
    add('MAG_SEQ_NEG_V',  0x2604)
    add('ANG_SEQ_POS_V',  0x2606)
    add('ANG_SEQ_ZERO_V', 0x2608)
    add('ANG_SEQ_NEG_V',  0x260A)
    add('UNBL_V_%',       0x260C)

    # 各相电压谐波（每相 5 指标 + 30 次谐波占 60 寄存器，步长 70）
    for phase, base in [('a', 0x260E), ('b', 0x2654), ('c', 0x269A)]:
        add(f'THD_V_{phase}_%',  base + 0)
        add(f'OTHD_V_{phase}_%', base + 2)
        add(f'ETHD_V_{phase}_%', base + 4)
        add(f'CF_V_{phase}_%',   base + 6)
        add(f'THFF_V_{phase}_%', base + 8)

    add('THD_V_avg_%',  0x26E0)
    add('OTHD_V_avg_%', 0x26E2)
    add('ETHD_V_avg_%', 0x26E4)

    # 输入通道电流谐波 (0x26E6, 每通道 68 寄存器)
    for ch in range(1, 25):
        n = f'{ch:03d}'
        base = 0x26E6 + (ch - 1) * 68
        add(f'THD_I_{n}_%',  base + 0)
        add(f'OTHD_I_{n}_%', base + 2)
        add(f'ETHD_I_{n}_%', base + 4)
        add(f'KF_I_{n}_%',   base + 6)

    # 用户通道电流序分量 (0x2D46, 每通道 14 寄存器)
    for ch in range(1, 13):
        n = f's{ch:03d}'
        base = 0x2D46 + (ch - 1) * 14
        add(f'MAG_SEQ_POS_I_{n}',  base + 0)
        add(f'MAG_SEQ_ZERO_I_{n}', base + 2)
        add(f'MAG_SEQ_NEG_I_{n}',  base + 4)
        add(f'ANG_SEQ_POS_I_{n}',  base + 6)
        add(f'ANG_SEQ_ZERO_I_{n}', base + 8)
        add(f'ANG_SEQ_NEG_I_{n}',  base + 10)
        add(f'UNBL_I_{n}_%',        base + 12)

    # ── 各相电压逐次谐波 HAR2-HAR31 ─────────────────────────────────────────
    # Phase A: 0x2618, Phase B: 0x265E, Phase C: 0x26A4（紧接 THFF 之后）
    for phase, har_base in [('a', 0x2618), ('b', 0x265E), ('c', 0x26A4)]:
        for k in range(2, 32):
            add(f'HAR{k}_V_{phase}_%', har_base + (k - 2) * 2)

    # ── 输入通道电流逐次谐波 HAR2-HAR31 (每通道块 68 寄存器，谐波在 +8 处) ──
    # 通道块基址: 0x26E6 + (ch-1)*68；THD/OTHD/ETHD/KF 占前 8 寄存器
    for ch in range(1, 25):
        n = f'{ch:03d}'
        ch_base = 0x26E6 + (ch - 1) * 68
        for k in range(2, 32):
            add(f'HAR{k}_I_{n}_%', ch_base + 8 + (k - 2) * 2)

    # ── 电能 (0x3000, double = 4 寄存器) ────────────────────────────────────
    for phase, base in [('a', 0x3000), ('b', 0x3024), ('c', 0x3048)]:
        add(f'EP_IMP_{phase}_kWh',     base + 0,  F64)
        add(f'EP_EXP_{phase}_kWh',     base + 4,  F64)
        add(f'EP_NET_{phase}_kWh',     base + 8,  F64)
        add(f'EP_TOTAL_{phase}_kWh',   base + 12, F64)
        add(f'EQ_IMP_{phase}_kvarh',   base + 16, F64)
        add(f'EQ_EXP_{phase}_kvarh',   base + 20, F64)
        add(f'EQ_NET_{phase}_kvarh',   base + 24, F64)
        add(f'EQ_TOTAL_{phase}_kvarh', base + 28, F64)
        add(f'ES_{phase}_kVAh',        base + 32, F64)

    base = 0x306C
    for key, off in [
        ('EP_IMP_kWh', 0),     ('EP_EXP_kWh', 4),
        ('EP_NET_kWh', 8),     ('EP_TOTAL_kWh', 12),
        ('EQ_IMP_kvarh', 16),  ('EQ_EXP_kvarh', 20),
        ('EQ_NET_kvarh', 24),  ('EQ_TOTAL_kvarh', 28),
        ('ES_kVAh', 32),
    ]:
        add(key, base + off, F64)

    # 输入通道电能 (0x3090, 每通道 36 寄存器)
    for ch in range(1, 25):
        n = f'{ch:03d}'
        base = 0x3090 + (ch - 1) * 36
        add(f'EP_IMP_{n}_kWh',     base + 0,  F64)
        add(f'EP_EXP_{n}_kWh',     base + 4,  F64)
        add(f'EP_NET_{n}_kWh',     base + 8,  F64)
        add(f'EP_TOTAL_{n}_kWh',   base + 12, F64)
        add(f'EQ_IMP_{n}_kvarh',   base + 16, F64)
        add(f'EQ_EXP_{n}_kvarh',   base + 20, F64)
        add(f'EQ_NET_{n}_kvarh',   base + 24, F64)
        add(f'EQ_TOTAL_{n}_kvarh', base + 28, F64)
        add(f'ES_{n}_kVAh',        base + 32, F64)

    # 用户通道电能 (0x33F0, 每通道 36 寄存器)
    for ch in range(1, 13):
        n = f's{ch:03d}'
        base = 0x33F0 + (ch - 1) * 36
        add(f'EP_IMP_{n}_kWh',     base + 0,  F64)
        add(f'EP_EXP_{n}_kWh',     base + 4,  F64)
        add(f'EP_NET_{n}_kWh',     base + 8,  F64)
        add(f'EP_TOTAL_{n}_kWh',   base + 12, F64)
        add(f'EQ_IMP_{n}_kvarh',   base + 16, F64)
        add(f'EQ_EXP_{n}_kvarh',   base + 20, F64)
        add(f'EQ_NET_{n}_kvarh',   base + 24, F64)
        add(f'EQ_TOTAL_{n}_kvarh', base + 28, F64)
        add(f'ES_{n}_kVAh',        base + 32, F64)

    # ── DI 状态 (FC02 离散输入 0x0000-0x0003) ────────────────────────────────
    for ch in range(1, 5):
        add(f'DI_ST_{ch:03d}', ch - 1, 'bit', fc=2)

    # ── DI 脉冲计数 (FC03 uint32, 0x6400 每通道 2 寄存器) ─────────────────────
    for ch in range(1, 5):
        add(f'DI_PC_{ch:03d}', 0x6400 + (ch - 1) * 2, 'uint32', fc=3)

    # ── DO 状态 (FC01 线圈 0x0000-0x0007) ────────────────────────────────────
    for ch in range(1, 9):
        add(f'DO_ST_{ch:03d}', ch - 1, 'bit', fc=1)

    # ── RO 状态 (FC01 线圈 0x0020-0x0021) ────────────────────────────────────
    for ch in range(1, 3):
        add(f'RO_ST_{ch:03d}', 0x0020 + (ch - 1), 'bit', fc=1)

    return regs


# ─────────────────────────────────────────────────────────────────────────────
# AcuCloud xlsx 列标题 → param_key 映射
# ─────────────────────────────────────────────────────────────────────────────

_DEG = '°'  # 度号，AcuCloud xlsx 中角度列单位的实际字符


def build_cloud_col_map() -> dict[str, str]:
    """返回 AcuCloud 导出 xlsx 的列标题 → param_key 映射（AcuRev4100）。"""
    m: dict[str, str] = {}

    # ── 系统基础量 ────────────────────────────────────────────────────────────
    m['System Frequency (Hz)']             = 'FREQ_Hz'
    m['Phase A Line-to-Neutral Voltage (V)']   = 'VLN_a_V'
    m['Phase B Line-to-Neutral Voltage (V)']   = 'VLN_b_V'
    m['Phase C Line-to-Neutral Voltage (V)']   = 'VLN_c_V'
    m['Average Line-to-Neutral Voltage (V)']   = 'VLN_avg_V'
    m['Phase A-B Line-to-Line Voltage (V)']    = 'VLL_ab_V'
    m['Phase B-C Line-to-Line Voltage (V)']    = 'VLL_bc_V'
    m['Phase C-A Line-to-Line Voltage (V)']    = 'VLL_ca_V'
    m['Average Line-to-Line Voltage (V)']      = 'VLL_avg_V'
    m[f'Phase A Line-to-Neutral Voltage Phase Angle ({_DEG})'] = 'ANG_VLN_a'
    m[f'Phase B Line-to-Neutral Voltage Phase Angle ({_DEG})'] = 'ANG_VLN_b'
    m[f'Phase C Line-to-Neutral Voltage Phase Angle ({_DEG})'] = 'ANG_VLN_c'

    # ── 各相 I/P/Q/S/PF ───────────────────────────────────────────────────────
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Current (A)']           = f'I_{_phase}_A'
        m[f'Phase {_ph} Active Power (kW)']      = f'P_{_phase}_kW'
        m[f'Phase {_ph} Reactive Power (kvar)']  = f'Q_{_phase}_kvar'
        m[f'Phase {_ph} Apparent Power (kVA)']   = f'S_{_phase}_kVA'
        m[f'Phase {_ph} Power Factor']            = f'PF_{_phase}'

    # ── 系统合计 ──────────────────────────────────────────────────────────────
    m['System Average Current (A)']    = 'I_avg_A'
    m['System Active Power (kW)']      = 'P_kW'
    m['System Reactive Power (kvar)']  = 'Q_kvar'
    m['System Apparent Power (kVA)']   = 'S_kVA'
    m['System Power Factor']           = 'PF'

    # ── 输入通道 001-024 ──────────────────────────────────────────────────────
    for _ch in range(1, 25):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} Current (A)']                       = f'I_{_n}_A'
        m[f'Input Channel {_ch} Active Power (kW)']                  = f'P_{_n}_kW'
        m[f'Input Channel {_ch} Reactive Power (kvar)']              = f'Q_{_n}_kvar'
        m[f'Input Channel {_ch} Apparent Power (kVA)']               = f'S_{_n}_kVA'
        m[f'Input Channel {_ch} Power Factor']                        = f'PF_{_n}'
        m[f'Input Channel {_ch} Current Phase Angle ({_DEG})']       = f'ANG_I_{_n}'

    # ── 用户通道 s001-s012 ────────────────────────────────────────────────────
    for _ch in range(1, 13):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Current (A)']          = f'I_{_n}_A'
        m[f'User Channel {_ch} Active Power (kW)']     = f'P_{_n}_kW'
        m[f'User Channel {_ch} Reactive Power (kvar)'] = f'Q_{_n}_kvar'
        m[f'User Channel {_ch} Apparent Power (kVA)']  = f'S_{_n}_kVA'
        m[f'User Channel {_ch} Power Factor']           = f'PF_{_n}'

    # ── 相需量 ────────────────────────────────────────────────────────────────
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Current Demand (A)']                    = f'DMD_I_{_phase}_A'
        m[f'Phase {_ph} Import Active Power Demand (kW)']        = f'DMD_IMP_P_{_phase}_kW'
        m[f'Phase {_ph} Export Active Power Demand (kW)']        = f'DMD_EXP_P_{_phase}_kW'
        m[f'Phase {_ph} Import Reactive Power Demand (kvar)']    = f'DMD_IMP_Q_{_phase}_kvar'
        m[f'Phase {_ph} Export Reactive Power Demand (kvar)']    = f'DMD_EXP_Q_{_phase}_kvar'
        m[f'Phase {_ph} Apparent Power Demand (kVA)']            = f'DMD_S_{_phase}_kVA'

    # ── 系统需量 ──────────────────────────────────────────────────────────────
    m['System Average Current Demand (A)']           = 'DMD_I_avg_A'
    m['System Import Active Power Demand (kW)']       = 'DMD_IMP_P_kW'
    m['System Export Active Power Demand (kW)']       = 'DMD_EXP_P_kW'
    m['System Import Reactive Power Demand (kvar)']   = 'DMD_IMP_Q_kvar'
    m['System Export Reactive Power Demand (kvar)']   = 'DMD_EXP_Q_kvar'
    m['System Apparent Power Demand (kVA)']           = 'DMD_S_kVA'

    # ── 输入通道需量 001-024 ──────────────────────────────────────────────────
    for _ch in range(1, 25):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} Current Demand (A)']                 = f'DMD_I_{_n}_A'
        m[f'Input Channel {_ch} Import Active Power Demand (kW)']     = f'DMD_IMP_P_{_n}_kW'
        m[f'Input Channel {_ch} Export Active Power Demand (kW)']     = f'DMD_EXP_P_{_n}_kW'
        m[f'Input Channel {_ch} Import Reactive Power Demand (kvar)'] = f'DMD_IMP_Q_{_n}_kvar'
        m[f'Input Channel {_ch} Export Reactive Power Demand (kvar)'] = f'DMD_EXP_Q_{_n}_kvar'
        m[f'Input Channel {_ch} Apparent Power Demand (kVA)']         = f'DMD_S_{_n}_kVA'

    # ── 用户通道需量 s001-s012 ────────────────────────────────────────────────
    for _ch in range(1, 13):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Current Demand (A)']                 = f'DMD_I_{_n}_A'
        m[f'User Channel {_ch} Import Active Power Demand (kW)']     = f'DMD_IMP_P_{_n}_kW'
        m[f'User Channel {_ch} Export Active Power Demand (kW)']     = f'DMD_EXP_P_{_n}_kW'
        m[f'User Channel {_ch} Import Reactive Power Demand (kvar)'] = f'DMD_IMP_Q_{_n}_kvar'
        m[f'User Channel {_ch} Export Reactive Power Demand (kvar)'] = f'DMD_EXP_Q_{_n}_kvar'
        m[f'User Channel {_ch} Apparent Power Demand (kVA)']         = f'DMD_S_{_n}_kVA'

    # ── 电能质量 ──────────────────────────────────────────────────────────────
    m['Voltage Positive Sequence Magnitude (V)']    = 'MAG_SEQ_POS_V'
    m['Voltage Zero Sequence Magnitude (V)']        = 'MAG_SEQ_ZERO_V'
    m['Voltage Negative Sequence Magnitude (V)']    = 'MAG_SEQ_NEG_V'
    m[f'Voltage Positive Sequence Angle ({_DEG})']  = 'ANG_SEQ_POS_V'
    m[f'Voltage Zero Sequence Angle ({_DEG})']      = 'ANG_SEQ_ZERO_V'
    m[f'Voltage Negative Sequence Angle ({_DEG})']  = 'ANG_SEQ_NEG_V'
    m['Voltage Unbalance Factor Magnitude (%)']     = 'UNBL_V_%'

    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Voltage THD (%)']         = f'THD_V_{_phase}_%'
        m[f'Phase {_ph} Voltage THD Odd (%)']      = f'OTHD_V_{_phase}_%'
        m[f'Phase {_ph} Voltage THD Even (%)']     = f'ETHD_V_{_phase}_%'
        m[f'Phase {_ph} Voltage Crest Factor (%)'] = f'CF_V_{_phase}_%'
        m[f'Phase {_ph} Voltage THFF Factor (%)']  = f'THFF_V_{_phase}_%'

    m['Average Voltage THD (%)']      = 'THD_V_avg_%'
    m['Average Voltage THD Odd (%)']  = 'OTHD_V_avg_%'
    m['Average Voltage THD Even (%)'] = 'ETHD_V_avg_%'

    # ── 输入通道谐波 001-024 ──────────────────────────────────────────────────
    for _ch in range(1, 25):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} Current THD (%)']       = f'THD_I_{_n}_%'
        m[f'Input Channel {_ch} Current THD Odd (%)']    = f'OTHD_I_{_n}_%'
        m[f'Input Channel {_ch} Current THD Even (%)']   = f'ETHD_I_{_n}_%'
        m[f'Input Channel {_ch} Current K Factor (%)']   = f'KF_I_{_n}_%'

    # ── 用户通道序分量 s001-s012 ──────────────────────────────────────────────
    for _ch in range(1, 13):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Current Positive Sequence Magnitude (A)']    = f'MAG_SEQ_POS_I_{_n}'
        m[f'User Channel {_ch} Current Zero Sequence Magnitude (A)']        = f'MAG_SEQ_ZERO_I_{_n}'
        m[f'User Channel {_ch} Current Negative Sequence Magnitude (A)']    = f'MAG_SEQ_NEG_I_{_n}'
        m[f'User Channel {_ch} Current Positive Sequence Angle ({_DEG})']   = f'ANG_SEQ_POS_I_{_n}'
        m[f'User Channel {_ch} Current Zero Sequence Angle ({_DEG})']       = f'ANG_SEQ_ZERO_I_{_n}'
        m[f'User Channel {_ch} Current Negative Sequence Angle ({_DEG})']   = f'ANG_SEQ_NEG_I_{_n}'
        m[f'User Channel {_ch} Current Unbalance Factor Magnitude (%)']     = f'UNBL_I_{_n}_%'

    # ── 相电能（首字母小写，与 xlsx 一致） ────────────────────────────────────
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} active energy import (kWh)']     = f'EP_IMP_{_phase}_kWh'
        m[f'Phase {_ph} active energy export (kWh)']     = f'EP_EXP_{_phase}_kWh'
        m[f'Phase {_ph} active energy net (kWh)']        = f'EP_NET_{_phase}_kWh'
        m[f'Phase {_ph} active energy total (kWh)']      = f'EP_TOTAL_{_phase}_kWh'
        m[f'Phase {_ph} reactive energy import (kvarh)'] = f'EQ_IMP_{_phase}_kvarh'
        m[f'Phase {_ph} reactive energy export (kvarh)'] = f'EQ_EXP_{_phase}_kvarh'
        m[f'Phase {_ph} reactive energy net (kvarh)']    = f'EQ_NET_{_phase}_kvarh'
        m[f'Phase {_ph} reactive energy total (kvarh)']  = f'EQ_TOTAL_{_phase}_kvarh'
        m[f'Phase {_ph} apparent energy (kVAh)']         = f'ES_{_phase}_kVAh'

    # ── 系统电能（首字母大写，与 xlsx 一致） ──────────────────────────────────
    m['System Import Active Energy (kWh)']    = 'EP_IMP_kWh'
    m['System Export Active Energy (kWh)']    = 'EP_EXP_kWh'
    m['System Net Active Energy (kWh)']       = 'EP_NET_kWh'
    m['System Total Active Energy (kWh)']     = 'EP_TOTAL_kWh'
    m['System Import Reactive Energy (kvarh)']= 'EQ_IMP_kvarh'
    m['System Export Reactive Energy (kvarh)']= 'EQ_EXP_kvarh'
    m['System Net Reactive Energy (kvarh)']   = 'EQ_NET_kvarh'
    m['System Total Reactive Energy (kvarh)'] = 'EQ_TOTAL_kvarh'
    m['System Apparent Energy (kVAh)']        = 'ES_kVAh'

    # ── 输入通道电能 001-024 ──────────────────────────────────────────────────
    for _ch in range(1, 25):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} active energy import (kWh)']     = f'EP_IMP_{_n}_kWh'
        m[f'Input Channel {_ch} active energy export (kWh)']     = f'EP_EXP_{_n}_kWh'
        m[f'Input Channel {_ch} active energy net (kWh)']        = f'EP_NET_{_n}_kWh'
        m[f'Input Channel {_ch} active energy total (kWh)']      = f'EP_TOTAL_{_n}_kWh'
        m[f'Input Channel {_ch} reactive energy import (kvarh)'] = f'EQ_IMP_{_n}_kvarh'
        m[f'Input Channel {_ch} reactive energy export (kvarh)'] = f'EQ_EXP_{_n}_kvarh'
        m[f'Input Channel {_ch} reactive energy net (kvarh)']    = f'EQ_NET_{_n}_kvarh'
        m[f'Input Channel {_ch} reactive energy total (kvarh)']  = f'EQ_TOTAL_{_n}_kvarh'
        m[f'Input Channel {_ch} apparent energy (kVAh)']         = f'ES_{_n}_kVAh'

    # ── 用户通道电能 s001-s012 ────────────────────────────────────────────────
    for _ch in range(1, 13):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} active energy import (kWh)']     = f'EP_IMP_{_n}_kWh'
        m[f'User Channel {_ch} active energy export (kWh)']     = f'EP_EXP_{_n}_kWh'
        m[f'User Channel {_ch} active energy net (kWh)']        = f'EP_NET_{_n}_kWh'
        m[f'User Channel {_ch} active energy total (kWh)']      = f'EP_TOTAL_{_n}_kWh'
        m[f'User Channel {_ch} reactive energy import (kvarh)'] = f'EQ_IMP_{_n}_kvarh'
        m[f'User Channel {_ch} reactive energy export (kvarh)'] = f'EQ_EXP_{_n}_kvarh'
        m[f'User Channel {_ch} reactive energy net (kvarh)']    = f'EQ_NET_{_n}_kvarh'
        m[f'User Channel {_ch} reactive energy total (kvarh)']  = f'EQ_TOTAL_{_n}_kvarh'
        m[f'User Channel {_ch} apparent energy (kVAh)']         = f'ES_{_n}_kVAh'

    # ── DI 脉冲 / 状态 ────────────────────────────────────────────────────────
    for _ch in range(1, 5):
        m[f'DI{_ch} pulse count'] = f'DI_PC_{_ch:03d}'
        m[f'DI{_ch} status']      = f'DI_ST_{_ch:03d}'

    # ── DO / RO 状态 ──────────────────────────────────────────────────────────
    for _ch in range(1, 9):
        m[f'DO{_ch} status'] = f'DO_ST_{_ch:03d}'
    for _ch in range(1, 3):
        m[f'RO{_ch} status'] = f'RO_ST_{_ch:03d}'

    return m
