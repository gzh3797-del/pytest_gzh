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

    voltage_wire_addr = 0x1042

    #  ================TOU Setting============================ #
    enable_tou = 0x601F
    monthly_billing_mode = 0x6020
    billing_time = 0x6021
    number_of_tariffs = 0x601B
    number_of_seasons = 0x6018
    number_of_schedules = 0x6019
    number_of_segments = 0x601A
    segment_1_setting = 0x604A

    weekend_schedule_ID = 0x601D

    dst_enable = 0x6000
    device_time = 0x1020

    enable_special_weekday_schedule = 0x668B
    holiday_setting_enable = 0x62F0
