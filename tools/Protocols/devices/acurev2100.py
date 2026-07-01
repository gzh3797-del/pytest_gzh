# -*- coding: utf-8 -*-
"""
devices/acurev2100.py — AcuRev2100 Modbus 参数地址映射

地址来源：AcuRev2100_ Modbus Address_v1.02_20260406.xlsx
FC 03H Read Holding Registers，float32 大端序，能量 uint32，PQ/谐波 uint16。

BACnet 对象名前缀：AcuRev2100-
对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """生成 AcuRev2100 param_key → ModbusRegister 映射。"""
    F32 = 'float32'
    U32 = 'uint32'
    regs: dict[str, ModbusRegister] = {}

    def add(key: str, addr: int, dtype: str = F32, scale: float = 1.0) -> None:
        regs[key] = ModbusRegister(key, addr, dtype, scale=scale)

    # ── 系统基础量 (0x2000) ───────────────────────────────────────────────────
    add('FREQ_Hz',    0x2000)
    add('VLN_a_V',    0x2002)
    add('VLN_b_V',    0x2004)
    add('VLN_c_V',    0x2006)
    add('VLN_avg_V',  0x2008)
    add('VLL_ab_V',   0x200A)
    add('VLL_bc_V',   0x200C)
    add('VLL_ca_V',   0x200E)
    add('VLL_avg_V',  0x2010)
    add('I_a_A',      0x2012)
    add('I_b_A',      0x2014)
    add('I_c_A',      0x2016)
    add('I_avg_A',    0x2018)
    add('P_kW',       0x201A)
    add('Q_kvar',     0x201C)
    add('S_kVA',      0x201E)
    add('PF',         0x2020)
    # 0x2022 = Load Nature S (float, 不在 EPICS 中，跳过)
    add('P_a_kW',     0x2024)
    add('P_b_kW',     0x2026)
    add('P_c_kW',     0x2028)
    add('Q_a_kvar',   0x202A)
    add('Q_b_kvar',   0x202C)
    add('Q_c_kvar',   0x202E)
    add('S_a_kVA',    0x2030)
    add('S_b_kVA',    0x2032)
    add('S_c_kVA',    0x2034)
    add('PF_a',       0x2036)
    add('PF_b',       0x2038)
    add('PF_c',       0x203A)
    # 0x203C/0x203E/0x2040 = Load Nature A/B/C (skipped)

    # ── 输入通道 001-018 (0x2100, 每通道 12 寄存器) ───────────────────────────
    # 内部顺序：I(2), P(2), Q(2), S(2), PF(2), LN(2, 跳过)
    for ch in range(1, 19):
        n = f'{ch:03d}'
        base = 0x2100 + (ch - 1) * 12
        add(f'I_{n}_A',    base + 0)
        add(f'P_{n}_kW',   base + 2)
        add(f'Q_{n}_kvar', base + 4)
        add(f'S_{n}_kVA',  base + 6)
        add(f'PF_{n}',     base + 8)

    # ── 用户通道 s001-s009 (0x21D8, 每通道 10 寄存器) ────────────────────────
    # 内部顺序：P(2), Q(2), S(2), PF(2), LN(2, 跳过)；无 I
    for ch in range(1, 10):
        n = f's{ch:03d}'
        base = 0x21D8 + (ch - 1) * 10
        add(f'P_{n}_kW',   base + 0)
        add(f'Q_{n}_kvar', base + 2)
        add(f'S_{n}_kVA',  base + 4)
        add(f'PF_{n}',     base + 6)

    # ── 相角 (0x3600, Word = uint16, 精度 0.1°, 范围 0~3600) ─────────────────
    # 0x3600: Vb相角, 0x3601: Vc相角, 0x3602~0x3613: I1~I18 相角，步长 1
    add('ANG_VLN_b', 0x3600, 'uint16', scale=0.1)
    add('ANG_VLN_c', 0x3601, 'uint16', scale=0.1)
    for ch in range(1, 19):
        add(f'ANG_I_{ch:03d}', 0x3602 + (ch - 1), 'uint16', scale=0.1)

    # ── 需量 (0x2D00, float32) ────────────────────────────────────────────────
    # 每指标块 = [当前需量(2), 预测(2), 最大值(2), 最大发生时间(3)] = 9 寄存器
    # 系统
    add('DMD_P_kW',       0x2D00)
    add('PRED_DMD_P_kW',  0x2D02)
    add('DMD_Q_kvar',     0x2D09)
    add('PRED_DMD_Q_kvar',0x2D0B)
    add('DMD_S_kVA',      0x2D12)
    add('PRED_DMD_S_kVA', 0x2D14)
    # 0x2D1B = 总电流需量（EPICS 里标相 A 电流需量，以下按相A处理）
    add('DMD_I_a_A',      0x2D1B)
    add('PRED_DMD_I_a_A', 0x2D1D)

    # 相 A 功率需量 (0x2D24)
    add('DMD_P_a_kW',       0x2D24)
    add('PRED_DMD_P_a_kW',  0x2D26)
    add('DMD_Q_a_kvar',     0x2D2D)
    add('PRED_DMD_Q_a_kvar',0x2D2F)
    add('DMD_S_a_kVA',      0x2D36)
    add('PRED_DMD_S_a_kVA', 0x2D38)

    # 相 B 功率需量 (0x2D3F)
    add('DMD_I_b_A',        0x2D3F)
    add('PRED_DMD_I_b_A',   0x2D41)
    add('DMD_P_b_kW',       0x2D48)
    add('PRED_DMD_P_b_kW',  0x2D4A)
    add('DMD_Q_b_kvar',     0x2D51)
    add('PRED_DMD_Q_b_kvar',0x2D53)
    add('DMD_S_b_kVA',      0x2D5A)
    add('PRED_DMD_S_b_kVA', 0x2D5C)

    # 相 C 功率需量 (0x2D63)
    add('DMD_I_c_A',        0x2D63)
    add('PRED_DMD_I_c_A',   0x2D65)
    add('DMD_P_c_kW',       0x2D6C)
    add('PRED_DMD_P_c_kW',  0x2D6E)
    add('DMD_Q_c_kvar',     0x2D75)
    add('PRED_DMD_Q_c_kvar',0x2D77)
    add('DMD_S_c_kVA',      0x2D7E)
    add('PRED_DMD_S_c_kVA', 0x2D80)

    # 输入通道需量 (0x2D87, 每通道 36 寄存器)
    # 内部顺序：I(9), P(9), Q(9), S(9)；每指标 = [DMD(2),PRED(2),max(2),time(3)]
    for ch in range(1, 19):
        n = f'{ch:03d}'
        base = 0x2D87 + (ch - 1) * 36
        add(f'DMD_I_{n}_A',        base + 0)
        add(f'PRED_DMD_I_{n}_A',   base + 2)
        add(f'DMD_P_{n}_kW',       base + 9)
        add(f'PRED_DMD_P_{n}_kW',  base + 11)
        add(f'DMD_Q_{n}_kvar',     base + 18)
        add(f'PRED_DMD_Q_{n}_kvar',base + 20)
        add(f'DMD_S_{n}_kVA',      base + 27)
        add(f'PRED_DMD_S_{n}_kVA', base + 29)

    # 用户通道需量 (0x300F, 每通道 27 寄存器，无 I)
    for ch in range(1, 10):
        n = f's{ch:03d}'
        base = 0x300F + (ch - 1) * 27
        add(f'DMD_P_{n}_kW',       base + 0)
        add(f'PRED_DMD_P_{n}_kW',  base + 2)
        add(f'DMD_Q_{n}_kvar',     base + 9)
        add(f'PRED_DMD_Q_{n}_kvar',base + 11)
        add(f'DMD_S_{n}_kVA',      base + 18)
        add(f'PRED_DMD_S_{n}_kVA', base + 20)

    # ── 有功电能 (0x2500, Dword = uint32, 精度 0.1 kWh) ─────────────────────
    E = 0.1  # 能量统一精度 0.1
    add('EP_IMP_a_kWh', 0x2500, U32, E)
    add('EP_IMP_b_kWh', 0x2502, U32, E)
    add('EP_IMP_c_kWh', 0x2504, U32, E)
    add('EP_IMP_kWh',   0x2506, U32, E)
    for ch in range(1, 19):
        add(f'EP_IMP_{ch:03d}_kWh', 0x2508 + (ch - 1) * 2, U32, E)
    for ch in range(1, 10):
        add(f'EP_IMP_s{ch:03d}_kWh', 0x252C + (ch - 1) * 2, U32, E)

    # ── 无功电能 (0x2B00, Dword = uint32, 精度 0.1 kvarh) ────────────────────
    add('EQ_IMP_a_kvarh', 0x2B00, U32, E)
    add('EQ_IMP_b_kvarh', 0x2B02, U32, E)
    add('EQ_IMP_c_kvarh', 0x2B04, U32, E)
    add('EQ_IMP_kvarh',   0x2B06, U32, E)
    for ch in range(1, 19):
        add(f'EQ_IMP_{ch:03d}_kvarh', 0x2B08 + (ch - 1) * 2, U32, E)
    for ch in range(1, 7):
        add(f'EQ_IMP_s{ch:03d}_kvarh', 0x2B2C + (ch - 1) * 2, U32, E)

    # ── 视在电能 (0x2B38, Dword = uint32, 精度 0.1 kVAh) ────────────────────
    add('ES_a_kVAh', 0x2B38, U32, E)
    add('ES_b_kVAh', 0x2B3A, U32, E)
    add('ES_c_kVAh', 0x2B3C, U32, E)
    add('ES_kVAh',   0x2B3E, U32, E)
    for ch in range(1, 19):
        add(f'ES_{ch:03d}_kVAh', 0x2B40 + (ch - 1) * 2, U32, E)
    for ch in range(1, 7):
        add(f'ES_s{ch:03d}_kVAh', 0x2B64 + (ch - 1) * 2, U32, E)

    # s007-s009 EQ 和 ES
    for ch in range(7, 10):
        add(f'EQ_IMP_s{ch:03d}_kvarh', 0x2B70 + (ch - 7) * 2, U32, E)
        add(f'ES_s{ch:03d}_kVAh',      0x2B76 + (ch - 7) * 2, U32, E)

    # ── PQ 电能质量 (0x3200, Word = uint16) ───────────────────────────────────
    # 缩放：CF 精度 0.001（÷1000），其余均为 0.01%（÷100）
    U16 = 'uint16'
    S_PCT  = 0.01   # THD / OTHD / ETHD / UNBL_I / THFF：0.01% 精度
    S_PCT1 = 0.1    # UNBL_V / KF：0.1% 精度
    S_CF   = 0.001  # Crest Factor：0.001 精度

    def add_pq(key: str, addr: int, scale: float = S_PCT) -> None:
        regs[key] = ModbusRegister(key, addr, U16, fc=3, scale=scale)

    # 电压不平衡 + 系统 THD
    add_pq('UNBL_V_%',    0x3200, S_PCT1)
    add_pq('THD_V_a_%',   0x3201)
    add_pq('THD_V_b_%',   0x3202)
    add_pq('THD_V_c_%',   0x3203)
    add_pq('THD_V_avg_%', 0x3204)
    # ── 电压逐次谐波 HAR2-HAR31（uint16, scale=0.01%）─────────────────────────
    for k in range(2, 32):
        add_pq(f'HAR{k}_V_a_%', 0x3205 + (k - 2))
        add_pq(f'HAR{k}_V_b_%', 0x3223 + (k - 2))
        add_pq(f'HAR{k}_V_c_%', 0x3241 + (k - 2))

    # 电流不平衡（0.01% 精度，与 UNBL_V 不同；勿误用 S_PCT1，否则读数偏大 10 倍）
    add_pq('UNBL_I_%',    0x325F, S_PCT)
    # 电压 Odd/Even THD + Crest Factor + THF（每相 4 寄存器，连续）
    for i, ph in enumerate(['a', 'b', 'c']):
        base = 0x3260 + i * 4
        add_pq(f'OTHD_V_{ph}_%', base + 0)
        add_pq(f'ETHD_V_{ph}_%', base + 1)
        add_pq(f'CF_V_{ph}_%',   base + 2, S_CF)
        add_pq(f'THFF_V_{ph}_%', base + 3)
    # 0x326C~0x329F = 保留（跳过）
    # 通道 001-018 Odd/Even THD + K Factor（每通道 3 寄存器，步长 3）
    for ch in range(1, 19):
        n = f'{ch:03d}'
        base = 0x32A0 + (ch - 1) * 3
        add_pq(f'OTHD_I_{n}_%', base + 0)
        add_pq(f'ETHD_I_{n}_%', base + 1)
        add_pq(f'KF_I_{n}_%',   base + 2, S_PCT1)
    # 通道 001-018 总 THD + 逐次谐波 HAR2-HAR31（每通道块 31 寄存器）
    for ch in range(1, 19):
        n = f'{ch:03d}'
        blk = 0x3300 + (ch - 1) * 31
        add_pq(f'THD_I_{n}_%', blk)
        for k in range(2, 32):
            add_pq(f'HAR{k}_I_{n}_%', blk + (k - 1))

    # 用户通道 s001-s009 电流不平衡（连续，步长 1）
    for ch in range(1, 10):
        add_pq(f'UNBL_I_s{ch:03d}_%', 0x352E + (ch - 1), S_PCT)

    return regs


# ─────────────────────────────────────────────────────────────────────────────
# AcuCloud xlsx 列标题 → param_key 映射
# ─────────────────────────────────────────────────────────────────────────────

_DEG = '°'


def build_cloud_col_map() -> dict[str, str]:
    """返回 AcuCloud 导出 xlsx 的列标题 → param_key 映射（AcuRev2100）。
    仅包含本设备 xlsx 中有实际数据的列，共 338 条。
    """
    m: dict[str, str] = {}

    # ── 相电压（xlsx 无 VLN_avg，VLL_avg 用 System Average Line-to-Line） ────
    m['Phase A Line-to-Neutral Voltage (V)']     = 'VLN_a_V'
    m['Phase B Line-to-Neutral Voltage (V)']     = 'VLN_b_V'
    m['Phase C Line-to-Neutral Voltage (V)']     = 'VLN_c_V'
    m['Phase A-B Line-to-Line Voltage (V)']      = 'VLL_ab_V'
    m['Phase B-C Line-to-Line Voltage (V)']      = 'VLL_bc_V'
    m['Phase C-A Line-to-Line Voltage (V)']      = 'VLL_ca_V'
    m['System Average Line-to-Line Voltage (V)'] = 'VLL_avg_V'

    # ── 各相 I/P/Q/S/PF ───────────────────────────────────────────────────────
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Current (A)']          = f'I_{_phase}_A'
        m[f'Phase {_ph} Active Power (kW)']    = f'P_{_phase}_kW'
        m[f'Phase {_ph} Reactive Power (kvar)']= f'Q_{_phase}_kvar'
        m[f'Phase {_ph} Apparent Power (kVA)'] = f'S_{_phase}_kVA'
        m[f'Phase {_ph} Power Factor']         = f'PF_{_phase}'

    # ── 系统合计 ──────────────────────────────────────────────────────────────
    m['System Average Current (A)']    = 'I_avg_A'
    m['System Active Power (kW)']      = 'P_kW'
    m['System Reactive Power (kvar)']  = 'Q_kvar'
    m['System Apparent Power (kVA)']   = 'S_kVA'
    m['System Power Factor']           = 'PF'

    # ── 输入通道 001-018 基础量 ───────────────────────────────────────────────
    # ch11 无电流测量（无 CT），其余字段齐全
    for _ch in range(1, 19):
        _n = f'{_ch:03d}'
        if _ch != 11:
            m[f'Input Channel {_ch} Current (A)']          = f'I_{_n}_A'
        m[f'Input Channel {_ch} Active Power (kW)']        = f'P_{_n}_kW'
        m[f'Input Channel {_ch} Reactive Power (kvar)']    = f'Q_{_n}_kvar'
        m[f'Input Channel {_ch} Apparent Power (kVA)']     = f'S_{_n}_kVA'
        m[f'Input Channel {_ch} Power Factor']             = f'PF_{_n}'

    # ── 用户通道 s001-s006（xlsx 中 s007-s009 未导出） ────────────────────────
    for _ch in range(1, 7):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Active Power (kW)']     = f'P_{_n}_kW'
        m[f'User Channel {_ch} Reactive Power (kvar)'] = f'Q_{_n}_kvar'
        m[f'User Channel {_ch} Apparent Power (kVA)']  = f'S_{_n}_kVA'
        m[f'User Channel {_ch} Power Factor']          = f'PF_{_n}'

    # ── 相需量 ────────────────────────────────────────────────────────────────
    m['Phase A Current Demand (A)'] = 'DMD_I_a_A'
    m['Phase B Current Demand (A)'] = 'DMD_I_b_A'
    m['Phase C Current Demand (A)'] = 'DMD_I_c_A'
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Active Power Demand (kW)']     = f'DMD_P_{_phase}_kW'
        m[f'Phase {_ph} Reactive Power Demand (kvar)'] = f'DMD_Q_{_phase}_kvar'
    # xlsx 中 Phase B Apparent Power Demand 列缺失（导出为第二个 Phase C），只映射 A 和 C
    m['Phase A Apparent Power Demand (kVA)'] = 'DMD_S_a_kVA'
    m['Phase C Apparent Power Demand (kVA)'] = 'DMD_S_c_kVA'

    # ── 系统需量 ──────────────────────────────────────────────────────────────
    m['System Active Power Demand (kW)']    = 'DMD_P_kW'
    m['System Reactive Power Demand (kvar)']= 'DMD_Q_kvar'
    m['System Apparent Power Demand (kVA)'] = 'DMD_S_kVA'

    # ── 输入通道需量 001-018 ──────────────────────────────────────────────────
    # ch16 无功需量列头拼写为 kvarh（xlsx 导出异常），其余用 kvar
    for _ch in range(1, 19):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} Current Demand (A)']        = f'DMD_I_{_n}_A'
        m[f'Input Channel {_ch} Active Power Demand (kW)']  = f'DMD_P_{_n}_kW'
        _rq = 'kvarh' if _ch == 16 else 'kvar'
        m[f'Input Channel {_ch} Reactive Power Demand ({_rq})'] = f'DMD_Q_{_n}_kvar'
        m[f'Input Channel {_ch} Apparent Power Demand (kVA)']   = f'DMD_S_{_n}_kVA'

    # ── 用户通道需量 s001-s006 ────────────────────────────────────────────────
    for _ch in range(1, 7):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Active Power Demand (kW)']     = f'DMD_P_{_n}_kW'
        m[f'User Channel {_ch} Reactive Power Demand (kvar)'] = f'DMD_Q_{_n}_kvar'
        m[f'User Channel {_ch} Apparent Power Demand (kVA)']  = f'DMD_S_{_n}_kVA'

    # ── 相电能（AcuRev2100 只有 Import） ──────────────────────────────────────
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Import Active Energy (kWh)']    = f'EP_IMP_{_phase}_kWh'
        m[f'Phase {_ph} Import Reactive Energy (kvarh)']= f'EQ_IMP_{_phase}_kvarh'
        m[f'Phase {_ph} Apparent Energy (kVAh)']        = f'ES_{_phase}_kVAh'

    # ── 系统电能 ──────────────────────────────────────────────────────────────
    m['System Import Active Energy (kWh)']    = 'EP_IMP_kWh'
    m['System Import Reactive Energy (kvarh)']= 'EQ_IMP_kvarh'
    m['System Apparent Energy (kVAh)']        = 'ES_kVAh'

    # ── 输入通道电能 001-018 ──────────────────────────────────────────────────
    # ch12 视在电能列缺失（导出异常），其余齐全
    for _ch in range(1, 19):
        _n = f'{_ch:03d}'
        m[f'Input Channel {_ch} Import Active Energy (kWh)']    = f'EP_IMP_{_n}_kWh'
        m[f'Input Channel {_ch} Import Reactive Energy (kvarh)']= f'EQ_IMP_{_n}_kvarh'
        if _ch != 12:
            m[f'Input Channel {_ch} Apparent Energy (kVAh)']    = f'ES_{_n}_kVAh'

    # ── 用户通道电能 s001-s006 ────────────────────────────────────────────────
    for _ch in range(1, 7):
        _n = f's{_ch:03d}'
        m[f'User Channel {_ch} Import Active Energy (kWh)']    = f'EP_IMP_{_n}_kWh'
        m[f'User Channel {_ch} Import Reactive Energy (kvarh)']= f'EQ_IMP_{_n}_kvarh'
        m[f'User Channel {_ch} Apparent Energy (kVAh)']        = f'ES_{_n}_kVAh'

    # ── 电能质量 ──────────────────────────────────────────────────────────────
    m['System Voltage Unbalance Factor (%)'] = 'UNBL_V_%'
    m['System Current Unbalance Factor (%)'] = 'UNBL_I_%'
    for _phase in ('a', 'b', 'c'):
        _ph = _phase.upper()
        m[f'Phase {_ph} Voltage THD (%)']         = f'THD_V_{_phase}_%'
        m[f'Phase {_ph} Odd Voltage THD (%)']      = f'OTHD_V_{_phase}_%'
        m[f'Phase {_ph} Voltage Crest Factor (%)'] = f'CF_V_{_phase}_%'

    return m
