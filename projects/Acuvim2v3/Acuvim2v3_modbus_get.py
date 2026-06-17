#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:Acuvim2v3_modbus_get.py
功能描述:寄存器值获取与比对
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import logging
import time

from comm.modbus_rtu_tcp import *
from comm.source_control import *
from tools.log import Log
from tools.excel_operate import data_read
from modbus_config import modbus_config
import math
import cmath
import threading
import time

Log(str(__file__).split("\\")[-1])

ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def line_to_line_voltage_calculate(ua, ub, uc, va_angle, vb_angle, vc_angle):
    """
    获取线电压值
    :param ua:ua
    :param ub:ub
    :param uc:uc
    :param va_angle:va_angle
    :param vb_angle:vb_angle
    :param vc_angle:vc_angle
    :return: vab, vbc, vca
    """
    ret = []
    va_complex = complex(math.cos(va_angle * math.pi / 180) * ua, math.sin(va_angle * math.pi / 180) * ua)
    vb_complex = complex(math.cos(vb_angle * math.pi / 180) * ub, math.sin(vb_angle * math.pi / 180) * ub)
    vc_complex = complex(math.cos(vc_angle * math.pi / 180) * uc, math.sin(vc_angle * math.pi / 180) * uc)
    vab = va_complex - vb_complex
    vbc = vb_complex - vc_complex
    vca = vc_complex - va_complex
    vab = abs(vab)
    vbc = abs(vbc)
    vca = abs(vca)
    ret.extend([vab, vbc, vca])
    return ret


def Active_Power_calculate(voltage, current, voltage_angle, current_angle):
    """
    计算有功功率
    :param voltage:电压
    :param current:电流
    :param voltage_angle:电压相位角度
    :param current_angle:电流相位角度
    :return:
    """
    voltage_current_angle = voltage_angle - current_angle
    Active_Power = (voltage * current * math.cos(math.radians(voltage_current_angle)))
    Active_Power = Active_Power
    return Active_Power


def Reactive_Power_calculate(voltage, current, voltage_angle, current_angle):
    """
    计算无功功率
    :param voltage: 电压
    :param current: 电流
    :param voltage_angle: 电压相位角度
    :param current_angle: 电流相位角度
    :return: 无功功率
    """
    voltage_current_angle = voltage_angle - current_angle
    Reactive_Power = (voltage * current * math.sin(math.radians(voltage_current_angle)))
    Reactive_Power = Reactive_Power
    return Reactive_Power


def Apparent_Power_calculate(voltage, current):
    """
    计算视在功率
    :param voltage: 电压
    :param current: 电流
    :return: 视在功率
    """
    Apparent_Power = (voltage * current)
    Apparent_Power = Apparent_Power
    return Apparent_Power


def Power_Factor_calculate(voltage_angle, current_angle):
    """
    计算power factor
    :param voltage_angle: 电压相位角度
    :param current_angle: 电流相位角度
    :return: power factor
    """
    voltage_current_angle = voltage_angle - current_angle
    Power_Factor = math.cos(math.radians(voltage_current_angle))
    return Power_Factor


def set_service_configuration_by_current(addr=0x1004, value=0):
    """
    寄存器配置-电流接线方式
    :param addr: 电压接线方式寄存器地址
    :param value: 电压接线方式值
    :return: True:写入成功,False:写入失败
    """
    ret = ModbusClient.write_registers(address=addr, values=value, slave=1)
    if f'{(addr, 1)}' not in str(ret):
        logging.error('Set_Service_Configuration fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=addr, count=1, slave=1)
    if ret[0] == value:
        return True
    return False


def set_service_configuration_by_voltage(addr=0x1003, value=0):
    """
    寄存器配置-电压接线方式
    :param addr: 电压接线方式寄存器地址
    :param value: 电压接线方式值
    :return: True:写入成功,False:写入失败
    """
    ret = ModbusClient.write_registers(address=addr, values=value, slave=1)
    if f'{(addr, 1)}' not in str(ret):
        logging.error('Set_Service_Configuration fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=addr, count=1, slave=1)
    if ret[0] == value:
        return True
    return False


def read_phase_voltage(standard_value, address=0x3002, times=5, interval_time=0.1):
    """
    相电压测量值和精度
    :param standard_value: 相位角度测量值和精度
    :param address: 标准值
    :param times: 寄存器地址
    :param interval_time: 采样次数
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address, count=2, slave=1)
        # print(value)
        if address == 0x3002:
            logging.info('Phase_A_Voltage ret is:{}'.format(value))
        elif address == 0x3004:
            logging.info('Phase_B_Voltage ret is:{}'.format(value))
        elif address == 0x3006:
            logging.info('Phase_C_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_line_voltage(standard_value, address=0x300A, times=5, interval_time=0.1):
    """
    线电压测量值和精度
    :param standard_value: 相位角度测量值和精度
    :param address: 标准值
    :param times: 寄存器地址
    :param interval_time: 采样次数
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address, count=2, slave=1)
        # print(value)
        if address == 0x300A:
            logging.info('Line_AB_Voltage ret is:{}'.format(value))
        elif address == 0x300C:
            logging.info('Line_BC_Voltage ret is:{}'.format(value))
        elif address == 0x300E:
            logging.info('Line_CA_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_phase_voltage_angle(standard_value, address=0x0000, times=5, interval_time=0.1):
    """
    电压相位角度测量值和精度
    :param standard_value: 相位角度测量值和精度
    :param address: 标准值
    :param times: 寄存器地址
    :param interval_time: 采样次数
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = [0x00]
        if address == 0x0000:
            value = [0x00]
            logging.info('Phase_A_Voltage_Angle ret is:{}'.format(value))
        if address == 0x42A0:
            value = ModbusClient.read_measurement(address=address, count=1, slave=1)
            logging.info('Phase_B_Voltage_Angle ret is:{}'.format(value))
        elif address == 0x42A1:
            value = ModbusClient.read_measurement(address=address, count=1, slave=1)
            logging.info('Phase_C_Voltage_Angle ret is:{}'.format(value))
        # print(f"value:{value}")
        # reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        # hex_num = reg.replace('0x', '')
        # integer_num = int(hex_num, 16)
        # value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        value_measu = value[0] / 10
        val_list.append(value_measu)
    # ModbusClient.close()
    # 新增角度0判断
    if standard_value == 0:
        for j in range(len(val_list)):
            if 350 <= val_list[j] <= 360:
                val_list[j] = val_list[j] - 360
    val_list.sort()
    min_val = min(val_list)
    max_val = max(val_list)
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((min_val - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((max_val - standard_value) / standard_value), 5)
        if min_val < 0:
            min_val = min_val + 360
        if max_val < 0:
            max_val = max_val + 360
        if avg_val < 0:
            avg_val = avg_val + 360
    else:
        if avg_val == standard_value == min_val == max_val == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [min_val, min_val_accuracy], [max_val, max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_current_angle(standard_value, address=0x42A2, times=5, interval_time=0.1):
    """
    电流相位角度测量值和精度
    :param standard_value: 时间间隔
    :param address:
    :param times:
    :param interval_time:
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=1, slave=1)
        if address == 0x42A2:
            logging.info('Input_A_Current_Phase_Angle ret is:{}'.format(value))
        if address == 0x42A3:
            logging.info('Input_B_Current_Phase_Angle ret is:{}'.format(value))
        elif address == 0x42A4:
            logging.info('Input_C_Current_Phase_Angle ret is:{}'.format(value))
        # reg = hex(value[0]).replace('0x', '').zfill(4)
        # hex_num = reg.replace('0x', '')
        # integer_num = int(hex_num, 16)
        # value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        value_measu = value[0] / 10
        val_list.append(value_measu)
        logging.info(f"第{v}次,address:{hex(address)},value_measu:{value_measu}")
        # print(f"第{v}次,address:{hex(address)},value_measu:{value_measu}")
    # ModbusClient.close()
    # 新增角度0判断
    if standard_value == 0:
        for j in range(len(val_list)):
            if 350 <= val_list[j] <= 360:
                val_list[j] = val_list[j] - 360
    # 原模块
    val_list.sort()
    min_val = min(val_list)
    max_val = max(val_list)
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((min_val - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((max_val - standard_value) / standard_value), 5)
        if min_val < 0:
            min_val = min_val + 360
        if max_val < 0:
            max_val = max_val + 360
        if avg_val < 0:
            avg_val = avg_val + 360
    else:
        if avg_val == standard_value == min_val == max_val == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [min_val, min_val_accuracy], [max_val, max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_current(standard_value, address=0x3012, times=5, interval_time=0.1):
    """
    电流测量值和精度
    :param standard_value:
    :param address:
    :param times:
    :param interval_time:
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=2, slave=1)
        if address == 0x3012:
            logging.info('Input_Channel_A_Current ret is:{}'.format(value))
        elif address == 0x3014:
            logging.info('Input_Channel_B_Current ret is:{}'.format(value))
        elif address == 0x3016:
            logging.info('Input_Channel_C_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_sys_power(standard_value, address=0x3022, times=5, interval_time=0.1):
    """
    系统功率测量值和精度
    :param standard_value:
    :param address:
    :param times:
    :param interval_time:
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=2, slave=1)
        if address == 0x3022:
            logging.info('Sys_ABC_Active_Power ret is:{}'.format(value))
        elif address == 0x302A:
            logging.info('Sys_ABC_Reactive_Power ret is:{}'.format(value))
        elif address == 0x3032:
            logging.info('Sys_ABC_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_active_power(standard_value, address=0x301C, times=5, interval_time=0.1):
    """
    有功功率测量值和精度
    :param standard_value:标准值
    :param address:寄存器地址
    :param times: 采样次数
    :param interval_time:时间间隔
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=2, slave=1)
        if address == 0x301C:
            logging.info('Channel_A_Active_Power ret is:{}'.format(value))
        elif address == 0x301E:
            logging.info('Channel_B_Active_Power ret is:{}'.format(value))
        elif address == 0x3020:
            logging.info('Channel_C_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_reactive_power(standard_value, address=0x3024, times=5, interval_time=0.1):
    """
    无功功率测量值和精度
    :param standard_value: 标准值
    :param address: 寄存器地址
    :param times: 采样次数
    :param interval_time: 时间间隔
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=2, slave=1)
        if address == 0x3024:
            logging.info('Input_A_Reactive_Power ret is:{}'.format(value))
        elif address == 0x3026:
            logging.info('Input_B_Reactive_Power ret is:{}'.format(value))
        elif address == 0x3028:
            logging.info('Input_C_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


def read_input_apparent_power(standard_value, address=0x302C, times=5, interval_time=0.1):
    """
    视在功率测量值和精度
    :param standard_value:
    :param address:
    :param times:
    :param interval_time:
    :return: (最小值、最小精度，最大值、最大精度，平均值、平均精度)
    """
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(interval_time)
        value = ModbusClient.read_measurement(address=address, count=2, slave=1)
        if address == 0x302C:
            logging.info('Input_A_Apparent_Power ret is:{}'.format(value))
        elif address == 0x302E:
            logging.info('Input_B_Apparent_Power ret is:{}'.format(value))
        elif address == 0x3030:
            logging.info('Input_C_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    if standard_value != 0:
        avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
        min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
        max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    else:
        if avg_val == standard_value == val_list[0] == val_list[-1] == 0 or avg_val < 0.001:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
        else:
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
    return [val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]


if __name__ == '__main__':
    # 关源
    switch_device_screen_interface(inter=0x00)
    set_ac(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    switch_device_screen_interface(inter=0x00)
