# -*- coding: utf-8 -*-
"""
devices/acuvimiiw.py — AcuvimIIW Modbus 参数地址映射

地址来源：Acuvim IIW&IIR&CL&EL Modbus Address v1.27_Haibo Song_260323.xlsx
FC 03H Read Holding Registers。

实时量（100ms）：float32，地址 0x3000+；V/I/PF/Freq scale=1.0，P/Q/S scale=0.001（寄存器单位 W/var/VA）
不平衡/需量（1S）：float32，地址 0x403C+；Unbl scale=100.0，DMD_P/Q/S scale=0.001（单位 W/var/VA）
电能（1S）       ：uint32/int32，地址 0x4048+，scale=0.01
谐波 THD（3S）   ：int16，地址 0x405A+，scale=0.01（two decimal）
谐波 PQ 质量（3S）：int16，scale=0.01
序分量           ：int16，地址 0x4294+
相角             ：uint16，地址 0x42A0+，scale=0.1
分相电能（1S）   ：uint32，地址 0x4620+，scale=0.01

BACnet 对象名前缀：AcuvimIIW-
对外接口：
    build_param_map() -> dict[str, ModbusRegister]
"""

from __future__ import annotations
from modbus_reader import ModbusRegister


def build_param_map() -> dict[str, ModbusRegister]:
    """生成 AcuvimIIW param_key → ModbusRegister 映射（477 个参数）。"""
    F32 = 'float32'
    U32 = 'uint32'
    I32 = 'int32'
    I16 = 'int16'
    U16 = 'uint16'
    regs: dict[str, ModbusRegister] = {}

    def add(key: str, addr: int, dtype: str = F32, scale: float = 1.0) -> None:
        regs[key] = ModbusRegister(key, addr, dtype, scale=scale)

    # ── 实时量（100ms block, 0x3000, float32, scale=1.0） ─────────────────────
    add('FREQ_Hz',   0x3000)
    add('VLN_a_V',   0x3002)   # Ua_rms
    add('VLN_b_V',   0x3004)   # Ub_rms
    add('VLN_c_V',   0x3006)   # Uc_rms
    add('VLN_avg_V', 0x3008)   # Uvag_rms
    add('VLL_ab_V',  0x300A)   # Uab_rms
    add('VLL_bc_V',  0x300C)   # Ubc_rms
    add('VLL_ca_V',  0x300E)   # Uca_rms
    add('VLL_avg_V', 0x3010)   # Ulag_rms
    add('I_a_A',     0x3012)   # Ia_rms
    add('I_b_A',     0x3014)   # Ib_rms
    add('I_c_A',     0x3016)   # Ic_rms
    add('I_avg_A',   0x3018)   # Ivag_rms
    add('I_n_A',     0x301A)   # In_rms
    # float32 功率寄存器单位为 W，乘以 0.001 转换为 kW/kvar/kVA
    add('P_a_kW',    0x301C, F32, 0.001)   # Pa_rms
    add('P_b_kW',    0x301E, F32, 0.001)   # Pb_rms
    add('P_c_kW',    0x3020, F32, 0.001)   # Pc_rms
    add('P_kW',      0x3022, F32, 0.001)   # P_rms
    add('Q_a_kvar',  0x3024, F32, 0.001)   # Qa_rms
    add('Q_b_kvar',  0x3026, F32, 0.001)   # Qb_rms
    add('Q_c_kvar',  0x3028, F32, 0.001)   # Qc_rms
    add('Q_kvar',    0x302A, F32, 0.001)   # Q_rms
    add('S_a_kVA',   0x302C, F32, 0.001)   # Sa_rms
    add('S_b_kVA',   0x302E, F32, 0.001)   # Sb_rms
    add('S_c_kVA',   0x3030, F32, 0.001)   # Sc_rms
    add('S_kVA',     0x3032, F32, 0.001)   # S_rms
    add('PF_a',      0x3034)   # PFa_rms
    add('PF_b',      0x3036)   # PFb_rms
    add('PF_c',      0x3038)   # PFc_rms
    add('PF',        0x303A)   # PF_rms

    # ── 不平衡 + 需量（1S block, 0x403C+, float32） ───────────────────────────
    # 0x4040 = Rlc_val（跳过）
    # float32 存 per-unit（0~1.0），仪表 HMI 显示为 %（×100），scale=100
    add('UNBL_V_%',  0x403C, F32, 100.0)
    add('UNBL_I_%',  0x403E, F32, 100.0)
    add('DMD_P_kW',  0x4042, F32, 0.001)   # P_dema，单位 W → kW
    add('DMD_Q_kvar',0x4044, F32, 0.001)   # Q_dema，单位 var → kvar
    add('DMD_S_kVA', 0x4046, F32, 0.001)   # S_dema，单位 VA → kVA

    # ── 电能（1S block, 0x4048+, uint32/int32, scale=0.1 kWh） ──────────────
    # IIW 上位机实测确认 BACnet 值正确，Modbus 读数 ×10 = 实际值，scale=0.1
    E = 0.1
    add('EP_IMP_kWh',    0x4048, U32, E)   # Ep_imp
    add('EP_EXP_kWh',    0x404A, U32, E)   # Ep_exp
    add('EQ_IMP_kvarh',  0x404C, U32, E)   # Eq_imp
    add('EQ_EXP_kvarh',  0x404E, U32, E)   # Eq_exp
    add('EP_TOTAL_kWh',  0x4050, U32, E)   # Ep_total
    add('EP_NET_kWh',    0x4052, I32, E)   # Ep_net（有符号，可为负）
    add('EQ_TOTAL_kvarh',0x4054, U32, E)   # Eq_total
    add('EQ_NET_kvarh',  0x4056, I32, E)   # Eq_net（有符号）
    add('ES_kVAh',       0x4058, U32, E)   # Es

    # ── 谐波 THD（3S block, 0x405A+, int16, scale=0.01, "two decimal"） ───────
    S = 0.01
    add('THD_V_a_%',   0x405A, I16, S)   # THD_V1(V12)
    add('THD_V_b_%',   0x405B, I16, S)   # THD_V2(V31)
    add('THD_V_c_%',   0x405C, I16, S)   # THD_V3(V23)
    add('THD_V_avg_%', 0x405D, I16, S)   # THD_avg
    add('THD_I_a_%',   0x405E, I16, S)   # THD_I1
    add('THD_I_b_%',   0x405F, I16, S)   # THD_I2
    add('THD_I_c_%',   0x4060, I16, S)   # THD_I3
    add('THD_I_avg_%', 0x4061, I16, S)   # THD_Iavg

    # ── OTHD/ETHD/CF/THFF（int16, scale=0.01） ────────────────────────────────
    # V1(V12) 2-31st harmonics at 0x4062-0x407F（30 regs，跳过）
    add('OTHD_V_a_%', 0x4080, I16, S)
    add('ETHD_V_a_%', 0x4081, I16, S)
    add('CF_V_a_%',   0x4082, I16, 0.001)   # CF 三位小数，不同于 THD 的 two decimal
    add('THFF_V_a_%', 0x4083, I16, S)
    # V2(V31) 2-31st harmonics at 0x4084-0x40A1（30 regs，跳过）
    add('OTHD_V_b_%', 0x40A2, I16, S)
    add('ETHD_V_b_%', 0x40A3, I16, S)
    add('CF_V_b_%',   0x40A4, I16, 0.001)
    add('THFF_V_b_%', 0x40A5, I16, S)
    # V3(V23) 2-31st harmonics at 0x40A6-0x40C3（30 regs，跳过）
    add('OTHD_V_c_%', 0x40C4, I16, S)
    add('ETHD_V_c_%', 0x40C5, I16, S)
    add('CF_V_c_%',   0x40C6, I16, 0.001)
    add('THFF_V_c_%', 0x40C7, I16, S)
    # I1 2-31st harmonics at 0x40C8-0x40E5（30 regs，跳过）
    add('OTHD_I_a_%', 0x40E6, I16, S)
    add('ETHD_I_a_%', 0x40E7, I16, S)
    add('KF_I_a_%',   0x40E8, I16, 0.1)   # IIW 实测 BACnet=1, Modbus raw=10 → scale=0.1
    # I2 2-31st harmonics at 0x40E9-0x4106（30 regs，跳过）
    add('OTHD_I_b_%', 0x4107, I16, S)
    add('ETHD_I_b_%', 0x4108, I16, S)
    add('KF_I_b_%',   0x4109, I16, 0.1)
    # I3 2-31st harmonics at 0x410A-0x4127（30 regs，跳过）
    add('OTHD_I_c_%', 0x4128, I16, S)
    add('ETHD_I_c_%', 0x4129, I16, S)
    add('KF_I_c_%',   0x412A, I16, 0.1)

    # ── 序分量（int16, 0x4294+） ──────────────────────────────────────────────
    # 电压序分量：scale=0.1V；电流序分量：scale=0.001A（需实测验证）
    S_SV = 0.1
    S_SI = 0.001
    add('SEQ_POS_REAL_V', 0x4294, I16, S_SV)
    add('SEQ_POS_IMG_V',  0x4295, I16, S_SV)
    add('SEQ_NEG_REAL_V', 0x4296, I16, S_SV)
    add('SEQ_NEG_IMG_V',  0x4297, I16, S_SV)
    add('SEQ_ZERO_REAL_V',0x4298, I16, S_SV)
    add('SEQ_ZERO_IMG_V', 0x4299, I16, S_SV)
    add('SEQ_POS_REAL_I', 0x429A, I16, S_SI)
    add('SEQ_POS_IMG_I',  0x429B, I16, S_SI)
    add('SEQ_NEG_REAL_I', 0x429C, I16, S_SI)
    add('SEQ_NEG_IMG_I',  0x429D, I16, S_SI)
    add('SEQ_ZERO_REAL_I',0x429E, I16, S_SI)
    add('SEQ_ZERO_IMG_I', 0x429F, I16, S_SI)

    # ── 相角（uint16, 0x42A0+, scale=0.1°, 范围 0~3600 → 0°~360°） ──────────
    add('ANG_VLN_b', 0x42A0, U16, 0.1)   # Phase Angle of V2 to V1
    add('ANG_VLN_c', 0x42A1, U16, 0.1)   # Phase Angle of V3 to V1
    add('ANG_I_a',   0x42A2, U16, 0.1)   # Phase Angle of I1 to V1
    add('ANG_I_b',   0x42A3, U16, 0.1)   # Phase Angle of I2 to V1
    add('ANG_I_c',   0x42A4, U16, 0.1)   # Phase Angle of I3 to V1
    add('ANG_VLN_a', 0x42A5, U16, 0.1)   # Phase Angle of V1 to V1

    # ── 逐次谐波 HAR2-HAR31（int16, scale=0.01，紧接各相 THD 块内） ──────────
    # V1(V12) at 0x4062, V2(V31) at 0x4084, V3(V23) at 0x40A6
    # I1 at 0x40C8, I2 at 0x40E9, I3 at 0x410A；每次谐波 1 寄存器步进
    for k in range(2, 32):
        add(f'HAR{k}_V_a_%', 0x4062 + (k - 2), I16, S)
        add(f'HAR{k}_V_b_%', 0x4084 + (k - 2), I16, S)
        add(f'HAR{k}_V_c_%', 0x40A6 + (k - 2), I16, S)
        add(f'HAR{k}_I_a_%', 0x40C8 + (k - 2), I16, S)
        add(f'HAR{k}_I_b_%', 0x40E9 + (k - 2), I16, S)
        add(f'HAR{k}_I_c_%', 0x410A + (k - 2), I16, S)

    # ── 逐次谐波 HAR32-HAR63（uint16/word, scale=0.01） ──────────────────────
    # V1 at 0x4500, V2 at 0x4520, V3 at 0x4540
    # I1 at 0x4560, I2 at 0x4580, I3 at 0x45A0；每次谐波 1 寄存器步进
    for k in range(32, 64):
        add(f'HAR{k}_V_a_%', 0x4500 + (k - 32), U16, S)
        add(f'HAR{k}_V_b_%', 0x4520 + (k - 32), U16, S)
        add(f'HAR{k}_V_c_%', 0x4540 + (k - 32), U16, S)
        add(f'HAR{k}_I_a_%', 0x4560 + (k - 32), U16, S)
        add(f'HAR{k}_I_b_%', 0x4580 + (k - 32), U16, S)
        add(f'HAR{k}_I_c_%', 0x45A0 + (k - 32), U16, S)

    # ── 分相电能（1S, 0x4620+, uint32, scale=0.01） ───────────────────────────
    add('EP_IMP_a_kWh',   0x4620, U32, E)   # Epa_imp
    add('EP_EXP_a_kWh',   0x4622, U32, E)   # Epa_exp
    add('EP_IMP_b_kWh',   0x4624, U32, E)   # Epb_imp
    add('EP_EXP_b_kWh',   0x4626, U32, E)   # Epb_exp
    add('EP_IMP_c_kWh',   0x4628, U32, E)   # Epc_imp
    add('EP_EXP_c_kWh',   0x462A, U32, E)   # Epc_exp
    add('EQ_IMP_a_kvarh', 0x462C, U32, E)   # Eqa_imp
    add('EQ_EXP_a_kvarh', 0x462E, U32, E)   # Eqa_exp
    add('EQ_IMP_b_kvarh', 0x4630, U32, E)   # Eqb_imp
    add('EQ_EXP_b_kvarh', 0x4632, U32, E)   # Eqb_exp
    add('EQ_IMP_c_kvarh', 0x4634, U32, E)   # Eqc_imp
    add('EQ_EXP_c_kvarh', 0x4636, U32, E)   # Eqc_exp
    add('ES_a_kVAh',      0x4638, U32, E)   # Esa
    add('ES_b_kVAh',      0x463A, U32, E)   # Esb
    add('ES_c_kVAh',      0x463C, U32, E)   # Esc

    return regs


def build_cloud_col_map() -> dict[str, str]:
    """AcuvimIIW AcuCloud xlsx列头 → param_key 映射（角度列用 ° U+00B0）。"""
    _D = '°'   # degree symbol used in IIW xlsx angle headers
    return {
        # ── 基础实时量 ──────────────────────────────────────────────────────────
        'System Frequency (Hz)':                        'FREQ_Hz',
        'Phase A Line-to-Neutral Voltage (V)':          'VLN_a_V',
        'Phase B Line-to-Neutral Voltage (V)':          'VLN_b_V',
        'Phase C Line-to-Neutral Voltage (V)':          'VLN_c_V',
        'System Average Line-to-Neutral Voltage (V)':   'VLN_avg_V',
        'Phase A-B Line-to-Line Voltage (V)':           'VLL_ab_V',
        'Phase B-C Line-to-Line Voltage (V)':           'VLL_bc_V',
        'Phase C-A Line-to-Line Voltage (V)':           'VLL_ca_V',
        'System Average Line-to-Line Voltage (V)':      'VLL_avg_V',
        'Phase A Current (A)':                          'I_a_A',
        'Phase B Current (A)':                          'I_b_A',
        'Phase C Current (A)':                          'I_c_A',
        'System Average Current (A)':                   'I_avg_A',
        'System Neutral Current (A)':                   'I_n_A',
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
        'Phase A Power Factor':                         'PF_a',
        'Phase B Power Factor':                         'PF_b',
        'Phase C Power Factor':                         'PF_c',
        'System Power Factor':                          'PF',

        # ── 相角（xlsx 使用 ° U+00B0） ─────────────────────────────────────────
        f'Phase B Voltage Phase Angle ({_D})':          'ANG_VLN_b',
        f'Phase C Voltage Phase Angle ({_D})':          'ANG_VLN_c',
        f'Phase A Current Phase Angle ({_D})':          'ANG_I_a',
        f'Phase B Current Phase Angle ({_D})':          'ANG_I_b',
        f'Phase C Current Phase Angle ({_D})':          'ANG_I_c',

        # ── 不平衡 + 需量 ──────────────────────────────────────────────────────
        'System Voltage Unbalance Factor (%)':          'UNBL_V_%',
        'System Current Unbalance Factor (%)':          'UNBL_I_%',
        'System Active Power Demand (kW)':              'DMD_P_kW',
        'System Reactive Power Demand (kvar)':          'DMD_Q_kvar',
        'System Apparent Power Demand (kVA)':           'DMD_S_kVA',

        # ── 系统电能 ────────────────────────────────────────────────────────────
        'System Import Active Energy (kWh)':            'EP_IMP_kWh',
        'System Export Active Energy (kWh)':            'EP_EXP_kWh',
        'System Import Reactive Energy (kvarh)':        'EQ_IMP_kvarh',
        'System Export Reactive Energy (kvarh)':        'EQ_EXP_kvarh',
        'System Total Active Energy (kWh)':             'EP_TOTAL_kWh',
        'System Net Active Energy (kWh)':               'EP_NET_kWh',
        'System Total Reactive Energy (kvarh)':         'EQ_TOTAL_kvarh',
        'System Net Reactive Energy (kvarh)':           'EQ_NET_kvarh',
        'System Apparent Energy (kVAh)':                'ES_kVAh',

        # ── 分相电能 ────────────────────────────────────────────────────────────
        'Phase A Import Active Energy (kWh)':           'EP_IMP_a_kWh',
        'Phase A Export Active Energy (kWh)':           'EP_EXP_a_kWh',
        'Phase B Import Active Energy (kWh)':           'EP_IMP_b_kWh',
        'Phase B Export Active Energy (kWh)':           'EP_EXP_b_kWh',
        'Phase C Import Active Energy (kWh)':           'EP_IMP_c_kWh',
        'Phase C Export Active Energy (kWh)':           'EP_EXP_c_kWh',
        'Phase A Import Reactive Energy (kvarh)':       'EQ_IMP_a_kvarh',
        'Phase A Export Reactive Energy (kvarh)':       'EQ_EXP_a_kvarh',
        'Phase B Import Reactive Energy (kvarh)':       'EQ_IMP_b_kvarh',
        'Phase B Export Reactive Energy (kvarh)':       'EQ_EXP_b_kvarh',
        'Phase C Import Reactive Energy (kvarh)':       'EQ_IMP_c_kvarh',
        'Phase C Export Reactive Energy (kvarh)':       'EQ_EXP_c_kvarh',
        'Phase A Apparent Energy (kVAh)':               'ES_a_kVAh',
        'Phase B Apparent Energy (kVAh)':               'ES_b_kVAh',
        'Phase C Apparent Energy (kVAh)':               'ES_c_kVAh',

        # ── THD 电压 ────────────────────────────────────────────────────────────
        'Phase A Voltage THD (%)':                      'THD_V_a_%',
        'Phase B Voltage THD (%)':                      'THD_V_b_%',
        'Phase C Voltage THD (%)':                      'THD_V_c_%',
        'System Average Voltage THD (%)':               'THD_V_avg_%',
        'Phase A Odd Voltage THD (%)':                  'OTHD_V_a_%',
        'Phase B Odd Voltage THD (%)':                  'OTHD_V_b_%',
        'Phase C Odd Voltage THD (%)':                  'OTHD_V_c_%',
        'Phase A Even Voltage THD (%)':                 'ETHD_V_a_%',
        'Phase B Even Voltage THD (%)':                 'ETHD_V_b_%',
        'Phase C Even Voltage THD (%)':                 'ETHD_V_c_%',
        'Phase A Voltage Crest Factor (%)':             'CF_V_a_%',
        'Phase B Voltage Crest Factor (%)':             'CF_V_b_%',
        'Phase C Voltage Crest Factor (%)':             'CF_V_c_%',
        'Phase A THFF (%)':                             'THFF_V_a_%',
        'Phase B THFF (%)':                             'THFF_V_b_%',
        'Phase C THFF (%)':                             'THFF_V_c_%',

        # ── THD 电流 ────────────────────────────────────────────────────────────
        'Phase A Current THD (%)':                      'THD_I_a_%',
        'Phase B Current THD (%)':                      'THD_I_b_%',
        'Phase C Current THD (%)':                      'THD_I_c_%',
        'System Average Current THD (%)':               'THD_I_avg_%',
        'Phase A Odd Current THD (%)':                  'OTHD_I_a_%',
        'Phase B Odd Current THD (%)':                  'OTHD_I_b_%',
        'Phase C Odd Current THD (%)':                  'OTHD_I_c_%',
        'Phase A Even Current THD (%)':                 'ETHD_I_a_%',
        'Phase B Even Current THD (%)':                 'ETHD_I_b_%',
        'Phase C Even Current THD (%)':                 'ETHD_I_c_%',
        'Phase A Current K Factor (%)':                 'KF_I_a_%',
        'Phase B Current K Factor (%)':                 'KF_I_b_%',
        'Phase C Current K Factor (%)':                 'KF_I_c_%',

        # ── 序分量（实部/虚部） ─────────────────────────────────────────────────
        'System Voltage Real Positive Sequence':        'SEQ_POS_REAL_V',
        'System Voltage Imaginary Positive Sequence':   'SEQ_POS_IMG_V',
        'System Voltage Real Negative Sequence':        'SEQ_NEG_REAL_V',
        'System Voltage Imaginary Negative Sequence':   'SEQ_NEG_IMG_V',
        'System Voltage Real Zero Sequence':            'SEQ_ZERO_REAL_V',
        'System Voltage Imaginary Zero Sequence':       'SEQ_ZERO_IMG_V',
        'System Current Real Positive Sequence':        'SEQ_POS_REAL_I',
        'System Current Imaginary Positive Sequence':   'SEQ_POS_IMG_I',
        'System Current Real Negative Sequence':        'SEQ_NEG_REAL_I',
        'System Current Imaginary Negative Sequence':   'SEQ_NEG_IMG_I',
        'System Current Real Zero Sequence':            'SEQ_ZERO_REAL_I',
        'System Current Imaginary Zero Sequence':       'SEQ_ZERO_IMG_I',
    }
