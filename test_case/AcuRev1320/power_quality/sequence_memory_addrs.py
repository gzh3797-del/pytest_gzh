#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:sequence_memory_addrs.py
功能描述:定义寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""


class SlaveId:
    slave_id = 1


class MemoryReg:
    reg_single = 1
    reg_double = 2
    reg_uint32 = 2
    reg_int32 = 2
    reg_uint16 = 6
    reg_float32 = 2


class MemoryAddr:
    #  ================200ms update============================ #

    voltage_unbalance_negative_addr = 0x9065
    current_unbalance_negative_addr = 0x9075
    voltage_zero_sequence_addr = 0x9069
    voltage_positive_sequence_addr = 0x906D
    voltage_negative_sequence_addr = 0x9071
    current_zero_sequence_addr = 0x9079
    current_positive_sequence_addr = 0x907D
    current_negative_sequence_addr = 0x9081
    phase_order_addr = 0x1063

    #  ================Moving Average 20ms update============================ #

    voltage_wire_addr = 0x1042
