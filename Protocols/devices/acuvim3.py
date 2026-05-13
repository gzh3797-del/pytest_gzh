# -*- coding: utf-8 -*-
"""
devices/acuvim3.py — AcuVIM3 Modbus 参数地址映射

地址来源：Acuvim3 User Modbus Address Table v1.08_Hongjian Zhu_260203.xlsx
FC 03H Read Holding Registers。

实时量（200ms block）：float32，地址 0x2114+，scale=1.0
需量（Demand）       ：float32，地址 0x6100+/0x6198+/0x61A8+，scale=1.0
电能（Energy）       ：float64（Double），地址 0x65C0+，scale=1.0

BACnet 对象名前缀：AcuVIM3-
对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """生成 AcuVIM3 param_key → ModbusRegister 映射（154 个参数）。"""
    F32 = 'float32'
    F64 = 'float64'
    regs: dict[str, ModbusRegister] = {}

    def add(key: str, addr: int, dtype: str = F32) -> None:
        regs[key] = ModbusRegister(key, addr, dtype)

    # ── 实时量（200ms block, 0x2114+, float32, scale=1.0） ─────────────────────
    add('FREQ_Hz',    0x2114)
    add('VLN_a_V',   0x2116)
    add('VLN_b_V',   0x2118)
    add('VLN_c_V',   0x211A)
    add('VLN_avg_V', 0x211C)
    add('VLL_ab_V',  0x211E)
    add('VLL_bc_V',  0x2120)
    add('VLL_ca_V',  0x2122)
    add('VLL_avg_V', 0x2124)
    add('I_a_A',     0x2126)
    add('I_b_A',     0x2128)
    add('I_c_A',     0x212A)
    add('I_n_A',     0x212C)
    add('I_avg_A',   0x212E)
    add('P_a_kW',    0x2130)
    add('P_b_kW',    0x2132)
    add('P_c_kW',    0x2134)
    add('P_kW',      0x2136)
    add('Q_a_kvar',  0x2138)
    add('Q_b_kvar',  0x213A)
    add('Q_c_kvar',  0x213C)
    add('Q_kvar',    0x213E)
    add('S_a_kVA',   0x2140)
    add('S_b_kVA',   0x2142)
    add('S_c_kVA',   0x2144)
    add('S_kVA',     0x2146)
    # 0x2148-0x214E = LC a/b/c/sys（Load Nature，不在 EPICS 中，跳过）
    add('PF_a',      0x2150)
    add('PF_b',      0x2152)
    add('PF_c',      0x2154)
    add('PF',        0x2156)
    add('LEAD_PF_a', 0x2158)
    add('LEAD_PF_b', 0x215A)
    add('LEAD_PF_c', 0x215C)
    add('LEAD_PF',   0x215E)
    add('LAG_PF_a',  0x2160)
    add('LAG_PF_b',  0x2162)
    add('LAG_PF_c',  0x2164)
    add('LAG_PF',    0x2166)

    # ── 相角（200ms block, float32, 单位 °） ──────────────────────────────────
    add('ANG_VLN_a', 0x2168)   # PA_VLN_a
    add('ANG_VLN_b', 0x216A)   # PA_VLN_b
    add('ANG_VLN_c', 0x216C)   # PA_VLN_c
    add('ANG_VLL_ab',0x216E)   # PA_VLL_ab
    add('ANG_VLL_bc',0x2170)   # PA_VLL_bc
    add('ANG_VLL_ca',0x2172)   # PA_VLL_ca
    add('ANG_I_a',   0x2174)   # PA_I_a
    add('ANG_I_b',   0x2176)   # PA_I_b
    add('ANG_I_c',   0x2178)   # PA_I_c

    # ── 序分量幅值（200ms block, float32） ────────────────────────────────────
    add('MAG_SEQ_POS_V',     0x217A)
    add('MAG_SEQ_ZERO_V',    0x217C)
    add('MAG_SEQ_NEG_V',     0x217E)
    add('SEQ_ZERO_RATIO_V_%',0x2180)
    add('UNBL_V_%',          0x2182)
    add('MAG_SEQ_POS_I',     0x2184)
    add('MAG_SEQ_ZERO_I',    0x2186)
    add('MAG_SEQ_NEG_I',     0x2188)
    add('SEQ_ZERO_RATIO_I_%',0x218A)
    add('UNBL_I_%',          0x218C)

    # ── 序分量相角（200ms block, float32, 单位 °） ────────────────────────────
    add('ANG_SEQ_POS_V',  0x218E)
    add('ANG_SEQ_ZERO_V', 0x2190)
    add('ANG_SEQ_NEG_V',  0x2192)
    add('ANG_SEQ_POS_I',  0x2194)
    add('ANG_SEQ_ZERO_I', 0x2196)
    add('ANG_SEQ_NEG_I',  0x2198)

    # ── THD 电压（200ms block, float32, 单位 %） ──────────────────────────────
    add('THD_V_a_%',   0x219A)
    add('THD_V_b_%',   0x219C)
    add('THD_V_c_%',   0x219E)
    add('OTHD_V_a_%',  0x21A0)
    add('OTHD_V_b_%',  0x21A2)
    add('OTHD_V_c_%',  0x21A4)
    add('ETHD_V_a_%',  0x21A6)
    add('ETHD_V_b_%',  0x21A8)
    add('ETHD_V_c_%',  0x21AA)
    add('CF_V_a_%',    0x21AC)
    add('CF_V_b_%',    0x21AE)
    add('CF_V_c_%',    0x21B0)

    # ── THD 电流（200ms block, float32, 单位 %） ──────────────────────────────
    add('THD_I_a_%',   0x21B2)
    add('THD_I_b_%',   0x21B4)
    add('THD_I_c_%',   0x21B6)
    add('THD_I_n_%',   0x21B8)
    add('OTHD_I_a_%',  0x21BA)
    add('OTHD_I_b_%',  0x21BC)
    add('OTHD_I_c_%',  0x21BE)
    add('OTHD_I_n_%',  0x21C0)
    add('ETHD_I_a_%',  0x21C2)
    add('ETHD_I_b_%',  0x21C4)
    add('ETHD_I_c_%',  0x21C6)
    add('ETHD_I_n_%',  0x21C8)
    add('TDD_I_a_%',   0x21CA)
    add('TDD_I_b_%',   0x21CC)
    add('TDD_I_c_%',   0x21CE)
    add('TDD_I_n_%',   0x21D0)
    add('CF_I_a_%',    0x21D2)
    add('CF_I_b_%',    0x21D4)
    add('CF_I_c_%',    0x21D6)
    add('CF_I_n_%',    0x21D8)
    add('KF_I_a_%',    0x21DA)
    add('KF_I_b_%',    0x21DC)
    add('KF_I_c_%',    0x21DE)
    add('KF_I_n_%',    0x21E0)

    # ── 闪变（200ms block, float32） ──────────────────────────────────────────
    add('FLICK_V_a', 0x21E2)
    add('FLICK_V_b', 0x21E4)
    add('FLICK_V_c', 0x21E6)

    # ── 需量电流（Demand block, 0x6100+, float32） ────────────────────────────
    add('DMD_I_a_A',    0x6100)
    add('DMD_I_b_A',    0x6102)
    add('DMD_I_c_A',    0x6104)
    add('DMD_I_avg_A',  0x6106)

    # ── 需量功率（Total Demand, 0x6198+/0x61A8+, float32） ───────────────────
    add('DMD_P_a_kW',   0x6198)   # DMD_TOTAL_P_a
    add('DMD_P_b_kW',   0x619A)   # DMD_TOTAL_P_b
    add('DMD_P_c_kW',   0x619C)   # DMD_TOTAL_P_c
    add('DMD_P_kW',     0x619E)   # DMD_TOTAL_P_sys
    add('DMD_Q_a_kvar', 0x61A0)   # DMD_TOTAL_Q_a
    add('DMD_Q_b_kvar', 0x61A2)   # DMD_TOTAL_Q_b
    add('DMD_Q_c_kvar', 0x61A4)   # DMD_TOTAL_Q_c
    add('DMD_Q_kvar',   0x61A6)   # DMD_TOTAL_Q_sys
    add('DMD_S_a_kVA',  0x61A8)
    add('DMD_S_b_kVA',  0x61AA)
    add('DMD_S_c_kVA',  0x61AC)
    add('DMD_S_kVA',    0x61AE)

    # ── 电能（Energy block, 0x65C0+, Double=float64, scale=1.0） ─────────────
    add('EP_IMP_a_kWh',    0x65C0, F64)
    add('EP_IMP_b_kWh',    0x65C4, F64)
    add('EP_IMP_c_kWh',    0x65C8, F64)
    add('EP_IMP_kWh',      0x65CC, F64)
    add('EQ_IMP_a_kvarh',  0x65D0, F64)
    add('EQ_IMP_b_kvarh',  0x65D4, F64)
    add('EQ_IMP_c_kvarh',  0x65D8, F64)
    add('EQ_IMP_kvarh',    0x65DC, F64)
    add('EP_EXP_a_kWh',    0x65E0, F64)
    add('EP_EXP_b_kWh',    0x65E4, F64)
    add('EP_EXP_c_kWh',    0x65E8, F64)
    add('EP_EXP_kWh',      0x65EC, F64)
    add('EQ_EXP_a_kvarh',  0x65F0, F64)
    add('EQ_EXP_b_kvarh',  0x65F4, F64)
    add('EQ_EXP_c_kvarh',  0x65F8, F64)
    add('EQ_EXP_kvarh',    0x65FC, F64)
    add('EP_NET_a_kWh',    0x6600, F64)
    add('EP_NET_b_kWh',    0x6604, F64)
    add('EP_NET_c_kWh',    0x6608, F64)
    add('EP_NET_kWh',      0x660C, F64)
    add('EQ_NET_a_kvarh',  0x6610, F64)
    add('EQ_NET_b_kvarh',  0x6614, F64)
    add('EQ_NET_c_kvarh',  0x6618, F64)
    add('EQ_NET_kvarh',    0x661C, F64)
    add('EP_TOTAL_a_kWh',  0x6620, F64)
    add('EP_TOTAL_b_kWh',  0x6624, F64)
    add('EP_TOTAL_c_kWh',  0x6628, F64)
    add('EP_TOTAL_kWh',    0x662C, F64)
    add('EQ_TOTAL_a_kvarh',0x6630, F64)
    add('EQ_TOTAL_b_kvarh',0x6634, F64)
    add('EQ_TOTAL_c_kvarh',0x6638, F64)
    add('EQ_TOTAL_kvarh',  0x663C, F64)
    add('ES_a_kVAh',       0x6640, F64)
    add('ES_b_kVAh',       0x6644, F64)
    add('ES_c_kVAh',       0x6648, F64)
    add('ES_kVAh',         0x664C, F64)

    return regs


def build_cloud_col_map() -> dict[str, str]:
    """AcuVIM3 AcuCloud xlsx列头 → param_key 映射（角度列用 (degree) 后缀）。"""
    return {
        # ── 基础实时量 ──────────────────────────────────────────────────────────
        'System Frequency (Hz)':                        'FREQ_Hz',
        'Line-to-Neutral Va (V)':                       'VLN_a_V',
        'Line-to-Neutral Vb (V)':                       'VLN_b_V',
        'Line-to-Neutral Vc (V)':                       'VLN_c_V',
        'System Average Line-to-Neutral Voltage (V)':   'VLN_avg_V',
        'Line to Line Vab (V)':                         'VLL_ab_V',
        'Line to Line Vbc (V)':                         'VLL_bc_V',
        'Line to Line Vca (V)':                         'VLL_ca_V',
        'Line to Line V system (V)':                    'VLL_avg_V',
        'Phase A Current (A)':                          'I_a_A',
        'Phase B Current (A)':                          'I_b_A',
        'Phase C Current (A)':                          'I_c_A',
        'System Neutral Current (A)':                   'I_n_A',
        'System Average Current (A)':                   'I_avg_A',
        'Phase A Active Power (kW)':                    'P_a_kW',
        'Phase B Active Power (kW)':                    'P_b_kW',
        'Phase C Active Power (kW)':                    'P_c_kW',
        'System Active Power (kW)':                     'P_kW',
        'Phase A Reactive Power (kvar)':                'Q_a_kvar',
        'Phase B Reactive Power (kvar)':                'Q_b_kvar',
        'Phase C Reactive Power (kvar)':                'Q_c_kvar',
        'System Reactive Power (kvar)':                 'Q_kvar',
        'Phase A Apparent Power (kVA)':                 'S_a_kVA',
        'Phase B Apparent Power (kVA)':                 'S_b_kVA',
        'Phase C Apparent Power (kVA)':                 'S_c_kVA',
        'System Apparent Power (kVA)':                  'S_kVA',

        # ── 相角（xlsx 使用 (degree) 文字） ────────────────────────────────────
        'Voltage AN Phase Angle (degree)':              'ANG_VLN_a',
        'Voltage BN Phase Angle (degree)':              'ANG_VLN_b',
        'Voltage CN Phase Angle (degree)':              'ANG_VLN_c',
        'Line Voltage AB Phase Angle (degree)':         'ANG_VLL_ab',
        'Line Voltage BC Phase Angle (degree)':         'ANG_VLL_bc',
        'Line Voltage CA Phase Angle (degree)':         'ANG_VLL_ca',
        'Phase A Current Phase Angle (degree)':         'ANG_I_a',
        'Phase B Current Phase Angle (degree)':         'ANG_I_b',
        'Phase C Current Phase Angle (degree)':         'ANG_I_c',

        # ── 序分量比率 + 不平衡 ────────────────────────────────────────────────
        'Voltage Zero Sequence Ratio (%)':              'SEQ_ZERO_RATIO_V_%',
        'Voltage Unbalance Factor (%)':                 'UNBL_V_%',
        'Current Zero Sequence Ratio (%)':              'SEQ_ZERO_RATIO_I_%',
        'Current Unbalance Factor (%)':                 'UNBL_I_%',

        # ── THD 电压 ────────────────────────────────────────────────────────────
        'Phase A Voltage THD (%)':                      'THD_V_a_%',
        'Phase B Voltage THD (%)':                      'THD_V_b_%',
        'Phase C Voltage THD (%)':                      'THD_V_c_%',
        'Phase A Odd Voltage THD (%)':                  'OTHD_V_a_%',
        'Phase B Odd Voltage THD (%)':                  'OTHD_V_b_%',
        'Phase C Odd Voltage THD (%)':                  'OTHD_V_c_%',
        'Phase A Even Voltage THD (%)':                 'ETHD_V_a_%',
        'Phase B Even Voltage THD (%)':                 'ETHD_V_b_%',
        'Phase C Even Voltage THD (%)':                 'ETHD_V_c_%',
        'Phase A Voltage Crest Factor (%)':             'CF_V_a_%',
        'Phase B Voltage Crest Factor (%)':             'CF_V_b_%',
        'Phase C Voltage Crest Factor (%)':             'CF_V_c_%',

        # ── THD 电流 ────────────────────────────────────────────────────────────
        'Phase A Current THD (%)':                      'THD_I_a_%',
        'Phase B Current THD (%)':                      'THD_I_b_%',
        'Phase C Current THD (%)':                      'THD_I_c_%',
        'Neutral Current Total Harmonic Distortion (%)': 'THD_I_n_%',
        'Phase A Odd Current THD (%)':                  'OTHD_I_a_%',
        'Phase B Odd Current THD (%)':                  'OTHD_I_b_%',
        'Phase C Odd Current THD (%)':                  'OTHD_I_c_%',
        'Neutral Current Odd-order Harmonic Distortion (%)': 'OTHD_I_n_%',
        'Phase A Even Current THD (%)':                 'ETHD_I_a_%',
        'Phase B Even Current THD (%)':                 'ETHD_I_b_%',
        'Phase C Even Current THD (%)':                 'ETHD_I_c_%',
        'Neutral Current Even-order Harmonic Distortion (%)': 'ETHD_I_n_%',
        'Phase A Current Total Demand Distortion (%)':  'TDD_I_a_%',
        'Phase B Current Total Demand Distortion (%)':  'TDD_I_b_%',
        'Phase C Current Total Demand Distortion (%)':  'TDD_I_c_%',
        'Neutral Current Total Demand Distortion (%)':  'TDD_I_n_%',
        'Phase A Current Crest Factor (No unit)':       'CF_I_a_%',
        'Phase B Current Crest Factor (No unit)':       'CF_I_b_%',
        'Phase C Current Crest Factor (No unit)':       'CF_I_c_%',
        'Neutral Current Crest Factor (No unit)':       'CF_I_n_%',
        'Phase A Current K Factor (%)':                 'KF_I_a_%',
        'Phase B Current K Factor (%)':                 'KF_I_b_%',
        'Phase C Current K Factor (%)':                 'KF_I_c_%',
        'Neutral Current K Factor (No unit)':           'KF_I_n_%',

        # ── 需量 ────────────────────────────────────────────────────────────────
        'Phase A Current Demand (A)':                   'DMD_I_a_A',
        'Phase B Current Demand (A)':                   'DMD_I_b_A',
        'Phase C Current Demand (A)':                   'DMD_I_c_A',
        'Average Current Demand (A)':                   'DMD_I_avg_A',
        'Phase A Active Power Demand (kW)':             'DMD_P_a_kW',
        'Phase B Active Power Demand (kW)':             'DMD_P_b_kW',
        'Phase C Active Power Demand (kW)':             'DMD_P_c_kW',
        'System Active Power Demand (kW)':              'DMD_P_kW',
        'Phase A Reactive Power Demand (kvar)':         'DMD_Q_a_kvar',
        'Phase B Reactive Power Demand (kvar)':         'DMD_Q_b_kvar',
        'Phase C Reactive Power Demand (kvar)':         'DMD_Q_c_kvar',
        'System Reactive Power Demand (kvar)':          'DMD_Q_kvar',
        'Phase A Apparent Power Demand (kVA)':          'DMD_S_a_kVA',
        'Phase C Apparent Power Demand (kVA)':          'DMD_S_c_kVA',
        'System Apparent Power Demand (kVA)':           'DMD_S_kVA',

        # ── 电能 ────────────────────────────────────────────────────────────────
        'Phase A Import Active Energy (kWh)':           'EP_IMP_a_kWh',
        'Phase B Import Active Energy (kWh)':           'EP_IMP_b_kWh',
        'Phase C Import Active Energy (kWh)':           'EP_IMP_c_kWh',
        'System Import Active Energy (kWh)':            'EP_IMP_kWh',
        'Phase A Reactive Import Energy (kvarh)':       'EQ_IMP_a_kvarh',
        'Phase B Reactive Import Energy (kvarh)':       'EQ_IMP_b_kvarh',
        'Phase C Reactive Import Energy (kvarh)':       'EQ_IMP_c_kvarh',
        'System Reactive Import Energy (kvarh)':        'EQ_IMP_kvarh',
        'Phase A Export Active Energy (kWh)':           'EP_EXP_a_kWh',
        'Phase B Export Active Energy (kWh)':           'EP_EXP_b_kWh',
        'Phase C Export Active Energy (kWh)':           'EP_EXP_c_kWh',
        'System Export Active Energy (kWh)':            'EP_EXP_kWh',
        'Phase A Export Reactive Energy (kvarh)':       'EQ_EXP_a_kvarh',
        'Phase B Export Reactive Energy (kvarh)':       'EQ_EXP_b_kvarh',
        'Phase C Export Reactive Energy (kvarh)':       'EQ_EXP_c_kvarh',
        'System Export Reactive Energy (kvarh)':        'EQ_EXP_kvarh',
        'Phase A Net Active Energy (kWh)':              'EP_NET_a_kWh',
        'Phase B Net Active Energy (kWh)':              'EP_NET_b_kWh',
        'Phase C Net Active Energy (kWh)':              'EP_NET_c_kWh',
        'System Net Active Energy (kWh)':               'EP_NET_kWh',
        'Phase A Net Reactive Energy (kvarh)':          'EQ_NET_a_kvarh',
        'Phase B Net Reactive Energy (kvarh)':          'EQ_NET_b_kvarh',
        'Phase C Net Reactive Energy (kvarh)':          'EQ_NET_c_kvarh',
        'System Net Reactive Energy (kvarh)':           'EQ_NET_kvarh',
        'Phase A Total Active Energy (kWh)':            'EP_TOTAL_a_kWh',
        'Phase B Total Active Energy (kWh)':            'EP_TOTAL_b_kWh',
        'Phase C Total Active Energy (kWh)':            'EP_TOTAL_c_kWh',
        'System Total Active Energy (kWh)':             'EP_TOTAL_kWh',
        'Phase A Total Reactive Energy (kvarh)':        'EQ_TOTAL_a_kvarh',
        'Phase B Total Reactive Energy (kvarh)':        'EQ_TOTAL_b_kvarh',
        'Phase C Total Reactive Energy (kvarh)':        'EQ_TOTAL_c_kvarh',
        'System Total Reactive Energy (kvarh)':         'EQ_TOTAL_kvarh',
        'Phase A Apparent Energy (kVAh)':               'ES_a_kVAh',
        'Phase B Apparent Energy (kVAh)':               'ES_b_kVAh',
        'Phase C Apparent Energy (kVAh)':               'ES_c_kVAh',
        'System Apparent Energy (kVAh)':                'ES_kVAh',
    }
