#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:memory_addrs.py
功能描述:定义寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""


class MemoryAddr:
    voltage_wire_addr = 0x1003
    current_wire_addr = 0x1004

    freq_rms_addr = 0x3000
    ua_rms_addr = 0x3002
    ub_rms_addr = 0x3004
    uc_rms_addr = 0x3006
    u_avg_rms_addr = 0x3008

    uab_rms_addr = 0x300A
    ubc_rms_addr = 0x300C
    uca_rms_addr = 0x300E
    ul_avg_rms_addr = 0x3010

    ia_rms_addr = 0x3012
    ib_rms_addr = 0x3014
    ic_rms_addr = 0x3016
    i_avg_rms_addr = 0x3018

    in_rms_addr = 0x301A
    pa_rms_addr = 0x301C
    pb_rms_addr = 0x301E
    pc_rms_addr = 0x3020
    qa_rms_addr = 0x3024
    qb_rms_addr = 0x3026
    qc_rms_addr = 0x3028
    sa_rms_addr = 0x302C
    sb_rms_addr = 0x302E
    sc_rms_addr = 0x3030
    p_total_rms_addr = 0x3022
    q_total_rms_addr = 0x302A
    s_total_rms_addr = 0x3032

    pf_a_rms_addr = 0x3034
    pf_b_rms_addr = 0x3036
    pf_c_rms_addr = 0x3038
    pf_total_rms_addr = 0x303A

    ua_p_rms_addr = 0x0000
    ub_p_rms_addr = 0x42A0
    uc_p_rms_addr = 0x42A1
    ia_p_rms_addr = 0x42A2
    ib_p_rms_addr = 0x42A3
    ic_p_rms_addr = 0x42A4
