#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @File     :1.py
# @Author   :lcs
# @Time     :2025/8/5
# @Desc     :

import time
from comm.modbus_rtu_tcp import *
from comm.source_control import *
from tools.log import Log
import xlwt
from tools.excel_operate import data_read
from modbus_config import modbus_config
import math
import cmath
import threading
import time
from tools.excel_operate import dcpara_4100addr_get

real_time_addr = dcpara_4100addr_get(r'/comm/test_data/AcuRev4100.xlsx', 'Real time')
basic_setting = dcpara_4100addr_get(r'/comm/test_data/AcuRev4100.xlsx', 'Basic Setting')
energy = dcpara_4100addr_get(r'/comm/test_data/AcuRev4100.xlsx', 'Energy')

Log(str(__file__).split("\\")[-1])

ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def read_frequency(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Frequency']['Start(Dec)'],
                                              count=real_time_addr['System Frequency']['Reg'], slave=1)
        logging.info('frequency ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Line-to-Neutral Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase A Line-to-Neutral Voltage']['Reg'], slave=1)
        logging.info('Phase_A_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Line-to-Neutral Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase B Line-to-Neutral Voltage']['Reg'], slave=1)
        logging.info('Phase_B_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Line-to-Neutral Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase C Line-to-Neutral Voltage']['Reg'], slave=1)
        logging.info('Phase_C_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_average_ln_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Average Line-to-Neutral Voltage']['Start(Dec)'],
                                              count=real_time_addr['Average Line-to-Neutral Voltage']['Reg'], slave=1)
        logging.info('Average_ln_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_ab_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase AB Line-to-Line Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase AB Line-to-Line Voltage']['Reg'], slave=1)
        logging.info('Phase_AB_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_bc_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase BC Line-to-Line Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase BC Line-to-Line Voltage']['Reg'], slave=1)
        logging.info('Phase_BC_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_ca_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase CA Line-to-Line Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase CA Line-to-Line Voltage']['Reg'], slave=1)
        logging.info('Phase_CA_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_average_ll_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Average Line-to-Line Voltage']['Start(Dec)'],
                                              count=real_time_addr['Average Line-to-Line Voltage']['Reg'], slave=1)
        logging.info('Average_ll_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_voltage_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr[
            'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
            'Start(Dec)'],
                                              count=real_time_addr[
                                                  'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
                                                  'Reg'], slave=1)
        logging.info('Phase_A_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_voltage_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr[
            'Phase B Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase BC Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
            'Start(Dec)'],
                                              count=real_time_addr[
                                                  'Phase B Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase BC Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
                                                  'Reg'], slave=1)
        logging.info('Phase_B_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_voltage_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr[
            'Phase C Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase CA Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
            'Start(Dec)'],
                                              count=real_time_addr[
                                                  'Phase C Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase CA Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
                                                  'Reg'], slave=1)
        logging.info('Phase_C_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Current']['Start(Dec)'],
                                              count=real_time_addr['Phase A Current']['Reg'], slave=1)
        logging.info('Phase_A_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8218, count=10, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
        read_list = []
        for i in range(5):
            if i != 3:
                i = i * 2
                reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
                hex_num = reg.replace('0x', '')
                integer_num = int(hex_num, 16)
                value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
                read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(4):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list
    # ModbusClient.close()


def read_phase_a_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Active Power']['Start(Dec)'],
                                              count=real_time_addr['Phase A Active Power']['Reg'], slave=1)
        logging.info('Phase_A_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Reactive Power']['Start(Dec)'],
                                              count=real_time_addr['Phase A Reactive Power']['Reg'], slave=1)
        logging.info('Phase_A_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Apparent Power']['Start(Dec)'],
                                              count=real_time_addr['Phase A Apparent Power']['Reg'], slave=1)
        logging.info('Phase_A_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_a_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=real_time_addr['Phase A Load Nature']['Start(Dec)'],
                                          count=real_time_addr['Phase A Load Nature']['Reg'], slave=1)
    logging.info('Phase_A_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_phase_a_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Power Factor']['Start(Dec)'],
                                              count=real_time_addr['Phase A Power Factor']['Reg'], slave=1)
        logging.info('Phase_A_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Current']['Start(Dec)'],
                                              count=real_time_addr['Phase B Current']['Reg'], slave=1)
        logging.info('Phase_B_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Active Power']['Start(Dec)'],
                                              count=real_time_addr['Phase B Active Power']['Reg'], slave=1)
        logging.info('Phase_B_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Reactive Power']['Start(Dec)'],
                                              count=real_time_addr['Phase B Reactive Power']['Reg'], slave=1)
        logging.info('Phase_B_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Apparent Power']['Start(Dec)'],
                                              count=real_time_addr['Phase B Apparent Power']['Reg'], slave=1)
        logging.info('Phase_B_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_b_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=real_time_addr['Phase B Load Nature']['Start(Dec)'],
                                          count=real_time_addr['Phase B Load Nature']['Reg'], slave=1)
    logging.info('Phase_B_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_phase_b_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase B Power Factor']['Start(Dec)'],
                                              count=real_time_addr['Phase B Power Factor']['Reg'], slave=1)
        logging.info('Phase_B_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Current']['Start(Dec)'],
                                              count=real_time_addr['Phase C Current']['Reg'], slave=1)
        logging.info('Phase_C_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Active Power']['Start(Dec)'],
                                              count=real_time_addr['Phase C Active Power']['Reg'], slave=1)
        logging.info('Phase_C_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Reactive Power']['Start(Dec)'],
                                              count=real_time_addr['Phase C Reactive Power']['Reg'], slave=1)
        logging.info('Phase_C_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Apparent Power']['Start(Dec)'],
                                              count=real_time_addr['Phase C Apparent Power']['Reg'], slave=1)
        logging.info('Phase_C_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_c_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=real_time_addr['Phase C Load Nature']['Start(Dec)'],
                                          count=real_time_addr['Phase C Load Nature']['Reg'], slave=1)
    logging.info('Phase_C_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_phase_c_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase C Power Factor']['Start(Dec)'],
                                              count=real_time_addr['Phase C Power Factor']['Reg'], slave=1)
        logging.info('Phase_C_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_system_average_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Average Current']['Start(Dec)'],
                                              count=real_time_addr['System Average Current']['Reg'], slave=1)
        logging.info('System_Average_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_system_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Active Power']['Start(Dec)'],
                                              count=real_time_addr['System Active Power']['Reg'], slave=1)
        logging.info('System_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_system_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Reactive Power']['Start(Dec)'],
                                              count=real_time_addr['System Reactive Power']['Reg'], slave=1)
        logging.info('System_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_system_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Apparent Power']['Start(Dec)'],
                                              count=real_time_addr['System Apparent Power']['Reg'], slave=1)
        logging.info('System_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_system_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=real_time_addr['System Load Nature']['Start(Dec)'],
                                          count=real_time_addr['System Load Nature']['Reg'], slave=1)
    logging.info('System_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_system_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['System Power Factor']['Start(Dec)'],
                                              count=real_time_addr['System Power Factor']['Reg'], slave=1)
        logging.info('System_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Current']['Start(Dec)'],
                                              count=real_time_addr['Input Channel 1 Current']['Reg'], slave=1)
        logging.info('Input_Channel_1_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Active Power']['Start(Dec)'],
                                              count=real_time_addr['Input Channel 1 Active Power']['Reg'], slave=1)
        logging.info('Channel_1_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Reactive Power']['Start(Dec)'],
                                              count=real_time_addr['Input Channel 1 Reactive Power']['Reg'], slave=1)
        logging.info('Input_Channel_1_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Apparent Power']['Start(Dec)'],
                                              count=real_time_addr['Input Channel 1 Apparent Power']['Reg'], slave=1)
        logging.info('Input_Channel_1_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Load Nature']['Start(Dec)'],
                                          count=real_time_addr['Input Channel 1 Load Nature']['Reg'], slave=1)
    logging.info('Input_Channel_1_Load_Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_input_channel_1_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Input Channel 1 Power Factor']['Start(Dec)'],
                                              count=real_time_addr['Input Channel 1 Power Factor']['Reg'], slave=1)
        logging.info('Input_Channel_1_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_1_current_phase_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(
            address=real_time_addr['Input Channel 1 Current Phase Angle']['Start(Dec)'],
            count=real_time_addr['Input Channel 1 Current Phase Angle']['Reg'], slave=1)
        logging.info('Input_Channel_1_Current_Phase_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_2_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Active Power']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Active Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Channel_2_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Reactive Power']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Reactive Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_2_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Apparent Power']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Apparent Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_2_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Load Nature']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Load Nature']['Reg']
    value = ModbusClient.read_measurement(address=address, count=count, slave=1)
    logging.info('Input_Channel_2_Load_Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_input_channel_2_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Power Factor']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Power Factor']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_2_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_2_current_phase_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current Phase Angle']['Start(Dec)'] + 14
    count = real_time_addr['Input Channel 1 Current Phase Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_2_Current_Phase_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_3_Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Active Power']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Active Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Channel_3_Active_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Reactive Power']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Reactive Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_3_Reactive_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Apparent Power']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Apparent Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_3_Apparent_Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Load Nature']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Load Nature']['Reg']
    value = ModbusClient.read_measurement(address=address, count=count, slave=1)
    logging.info('Input_Channel_3_Load_Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_input_channel_3_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Power Factor']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Power Factor']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_3_Power_Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_3_current_phase_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current Phase Angle']['Start(Dec)'] + 28
    count = real_time_addr['Input Channel 1 Current Phase Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input_Channel_3_Current_Phase_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Active Power']['Start(Dec)']
    count = real_time_addr['User Channel 1 Active Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Active Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Reactive Power']['Start(Dec)']
    count = real_time_addr['User Channel 1 Reactive Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Reactive Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Apparent Power']['Start(Dec)']
    count = real_time_addr['User Channel 1 Apparent Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Apparent Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Load Nature']['Start(Dec)']
    count = real_time_addr['User Channel 1 Load Nature']['Reg']
    value = ModbusClient.read_measurement(address=address, count=count, slave=1)
    logging.info('Input_Channel_3_Load_Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_user_channel_1_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Power Factor']['Start(Dec)']
    count = real_time_addr['User Channel 1 Power Factor']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Power Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_2_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 2 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_2_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Active Power']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Active Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 2 Active Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_2_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Reactive Power']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Reactive Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 2 Reactive Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_2_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Apparent Power']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Apparent Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 2 Apparent Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_2_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Load Nature']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Load Nature']['Reg']
    value = ModbusClient.read_measurement(address=address, count=count, slave=1)
    logging.info('User Channel 2 Load Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_user_channel_2_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Power Factor']['Start(Dec)'] + 12
    count = real_time_addr['User Channel 1 Power Factor']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 2 Power Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_3_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 3 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_3_active_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Active Power']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Active Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 3 Active Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_3_reactive_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Reactive Power']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Reactive Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 3 Reactive Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_3_apparent_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Apparent Power']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Apparent Power']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 3 Apparent Power ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_3_load_nature():
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Load Nature']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Load Nature']['Reg']
    value = ModbusClient.read_measurement(address=address, count=count, slave=1)
    logging.info('User Channel 3 Load Nature ret is:{}'.format(value))
    # ModbusClient.close()
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def read_user_channel_3_power_factor(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Power Factor']['Start(Dec)'] + 24
    count = real_time_addr['User Channel 1 Power Factor']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 3 Power Factor ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def line_to_line_voltage_calculate(ua: float, ub: float, uc: float, va_angle: float, vb_angle: float, vc_angle: float):
    """
    :param ua:
    :param ub:
    :param uc:
    :param va_angle:
    :param vb_angle:
    :param vc_angle:
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


def active_power_calculate(voltage, current, voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Active_Power = (voltage * current * math.cos(math.radians(voltage_current_angle))) / 1000
    Active_Power = Active_Power
    return Active_Power


def reactive_power_calculate(voltage, current, voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Reactive_Power = (voltage * current * math.sin(math.radians(voltage_current_angle))) / 1000
    Reactive_Power = Reactive_Power
    return Reactive_Power


def apparent_power_calculate(voltage, current):
    Apparent_Power = (voltage * current / 1000)
    Apparent_Power = Apparent_Power
    return Apparent_Power


def power_factor_calculate(voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Power_Factor = math.cos(math.radians(voltage_current_angle))
    return Power_Factor


def system_load_nature_calculate(P_Sum, Q_Sum):
    # 如果 P_Sum 和 Q_Sum 都是数字
    if isinstance(P_Sum, (int, float)) and isinstance(Q_Sum, (int, float)):
        if P_Sum > 0:
            if Q_Sum > 0:
                return "L"
            elif Q_Sum < 0:
                return "C"
            else:
                return "R"
        elif P_Sum < 0:
            if -0.00001 < P_Sum < 0 and Q_Sum < 0:
                return "C"
            elif Q_Sum > 0:
                return "C"
            else:
                return "L"
        else:
            if 0 <= P_Sum < 0.00001 and Q_Sum > 0:
                return "L"
            elif Q_Sum < 0.00001:
                return "C"
            else:
                return ""
    # 如果 P_Sum 是复数且 Q_Sum 小于 0.00001
    elif isinstance(P_Sum, complex) and abs(P_Sum.imag) > 0 and Q_Sum < 0.00001:
        return "C"
    else:
        return ""


def load_nature_calculate(Vol_Angle, Cur_Angle):
    Vol_Cur_Angle = Vol_Angle - Cur_Angle
    # 标准化角度到 0 到 360 度之间
    Vol_Cur_Angle = (Vol_Cur_Angle + 360) % 360

    if Vol_Cur_Angle == 0 or Vol_Cur_Angle == 180:
        return "R"
    elif (0 < Vol_Cur_Angle <= 90) or (180 < Vol_Cur_Angle <= 270):
        return "L"
    elif (90 < Vol_Cur_Angle < 180) or (270 < Vol_Cur_Angle < 360):
        return "C"
    else:
        return ""


def power_standard_value_calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang):
    """
    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :return:
    Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P, Phase_C_Q,Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF
    """
    Power = []
    Phase_A_P = active_power_calculate(Va, Ia, Va_ang, Ia_ang)
    Phase_A_Q = reactive_power_calculate(Va, Ia, Va_ang, Ia_ang)
    Phase_A_S = apparent_power_calculate(Va, Ia)
    Phase_A_PF = power_factor_calculate(Va_ang, Ia_ang)

    Phase_B_P = active_power_calculate(Vb, Ib, Vb_ang, Ib_ang)
    Phase_B_Q = reactive_power_calculate(Vb, Ib, Vb_ang, Ib_ang)
    Phase_B_S = apparent_power_calculate(Vb, Ib)
    Phase_B_PF = power_factor_calculate(Vb_ang, Ib_ang)

    Phase_C_P = active_power_calculate(Vc, Ic, Vc_ang, Ic_ang)
    Phase_C_Q = reactive_power_calculate(Vc, Ic, Vc_ang, Ic_ang)
    Phase_C_S = apparent_power_calculate(Vc, Ic)
    Phase_C_PF = power_factor_calculate(Vc_ang, Ic_ang)

    Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
    Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
    Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend(
        [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P, Phase_C_Q,
         Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
    return Power


def read_acurev4100_power(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, User_Channel):
    """
    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param User_Channel:
    :return: [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [Input1_P, scale_Input1_P], [Input1_Q, scale_Input1_Q],
         [Input1_S, scale_Input1_S], [Input1_PF, scale_Input1_PF], [Input2_P, scale_Input2_P],
         [Input2_Q, scale_Input2_Q], [Input2_S, scale_Input2_S], [Input2_PF, scale_Input2_PF],
         [Input3_P, scale_Input3_P], [Input3_Q, scale_Input3_Q],
         [Input3_S, scale_Input3_S], [Input3_PF, scale_Input3_PF]~~~]
    """
    power_list = []
    Power = power_standard_value_calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang)
    Phase_A_P = read_phase_a_active_power(Power[0], times=40)
    Phase_A_Q = read_phase_a_reactive_power(Power[1], times=40)
    Phase_A_S = read_phase_a_apparent_power(Power[2], times=40)
    Phase_A_PF = read_phase_a_power_factor(Power[3], times=40)

    if Power[0] != 0:
        scale_Phase_A_P = abs((Phase_A_P - Power[0]) / Power[0])
    else:
        if Power[0] == Phase_A_P == 0:
            scale_Phase_A_P = 0
        else:
            scale_Phase_A_P = 'null'

    if Power[1] != 0:
        scale_Phase_A_Q = abs((Phase_A_Q - Power[1]) / Power[1])
    else:
        if Power[1] == Phase_A_Q == 0:
            scale_Phase_A_Q = 0
        else:
            scale_Phase_A_Q = 'null'

    if Power[2] != 0:
        scale_Phase_A_S = abs((Phase_A_S - Power[2]) / Power[2])
    else:
        if Power[2] == Phase_A_S == 0:
            scale_Phase_A_S = 0
        else:
            scale_Phase_A_S = 'null'

    if Power[3] != 0:
        scale_Phase_A_PF = abs((Phase_A_PF - Power[3]) / Power[3])
    else:
        if Power[3] == Phase_A_PF == 0:
            scale_Phase_A_PF = 0
        else:
            scale_Phase_A_PF = 'null'

    Phase_B_P = read_phase_b_active_power(Power[4], times=40)
    Phase_B_Q = read_phase_b_reactive_power(Power[5], times=40)
    Phase_B_S = read_phase_b_apparent_power(Power[6], times=40)
    Phase_B_PF = read_phase_b_power_factor(Power[7], times=40)

    if Power[4] != 0:
        scale_Phase_B_P = abs((Phase_B_P - Power[4]) / Power[4])
    else:
        if Power[4] == Phase_B_P == 0:
            scale_Phase_B_P = 0
        else:
            scale_Phase_B_P = 'null'

    if Power[5] != 0:
        scale_Phase_B_Q = abs((Phase_B_Q - Power[5]) / Power[5])
    else:
        if Power[5] == Phase_B_Q == 0:
            scale_Phase_B_Q = 0
        else:
            scale_Phase_B_Q = 'null'

    if Power[6] != 0:
        scale_Phase_B_S = abs((Phase_B_S - Power[6]) / Power[6])
    else:
        if Power[6] == Phase_B_S == 0:
            scale_Phase_B_S = 0
        else:
            scale_Phase_B_S = 'null'

    if Power[7] != 0:
        scale_Phase_B_PF = abs((Phase_B_PF - Power[7]) / Power[7])
    else:
        if Power[7] == Phase_B_PF == 0:
            scale_Phase_B_PF = 0
        else:
            scale_Phase_B_PF = 'null'

    Phase_C_P = read_phase_c_active_power(Power[8], times=40)
    Phase_C_Q = read_phase_c_reactive_power(Power[9], times=40)
    Phase_C_S = read_phase_c_apparent_power(Power[10], times=40)
    Phase_C_PF = read_phase_c_power_factor(Power[11], times=40)

    if Power[8] != 0:
        scale_Phase_C_P = abs((Phase_C_P - Power[8]) / Power[8])
    else:
        if Power[8] == Phase_C_P == 0:
            scale_Phase_C_P = 0
        else:
            scale_Phase_C_P = 'null'

    if Power[9] != 0:
        scale_Phase_C_Q = abs((Phase_C_Q - Power[9]) / Power[9])
    else:
        if Power[9] == Phase_C_Q == 0:
            scale_Phase_C_Q = 0
        else:
            scale_Phase_C_Q = 'null'

    if Power[10] != 0:
        scale_Phase_C_S = abs((Phase_C_S - Power[10]) / Power[10])
    else:
        if Power[10] == Phase_C_S == 0:
            scale_Phase_C_S = 0
        else:
            scale_Phase_C_S = 'null'

    if Power[11] != 0:
        scale_Phase_C_PF = abs((Phase_C_PF - Power[11]) / Power[11])
    else:
        if Power[11] == Phase_C_PF == 0:
            scale_Phase_C_PF = 0
        else:
            scale_Phase_C_PF = 'null'

    Sys_P = read_system_active_power(Power[12], times=40)
    Sys_Q = read_system_reactive_power(Power[13], times=40)
    Sys_S = read_system_apparent_power(Power[14], times=40)
    Sys_PF = read_system_power_factor(Power[15], times=40)

    if Power[12] != 0:
        scale_Sys_P = abs((Sys_P - Power[12]) / Power[12])
    else:
        if Power[12] == Sys_P == 0:
            scale_Sys_P = 0
        else:
            scale_Sys_P = 'null'

    if Power[13] != 0:
        scale_Sys_Q = abs((Sys_Q - Power[13]) / Power[13])
    else:
        if Power[13] == Sys_Q == 0:
            scale_Sys_Q = 0
        else:
            scale_Sys_Q = 'null'

    if Power[14] != 0:
        scale_Sys_S = abs((Sys_S - Power[14]) / Power[14])
    else:
        if Power[14] == Sys_S == 0:
            scale_Sys_S = 0
        else:
            scale_Sys_S = 'null'

    if Power[15] != 0:
        scale_Sys_PF = abs((Sys_PF - Power[15]) / Power[15])
    else:
        if Power[15] == Sys_PF == 0:
            scale_Sys_PF = 0
        else:
            scale_Sys_PF = 'null'

    Input1_P = read_input_channel_1_active_power(Power[0], times=40)
    Input1_Q = read_input_channel_1_reactive_power(Power[1], times=40)
    Input1_S = read_input_channel_1_apparent_power(Power[2], times=40)
    Input1_PF = read_input_channel_1_power_factor(Power[3], times=40)

    if Power[0] != 0:
        scale_Input1_P = abs((Input1_P - Power[0]) / Power[0])
    else:
        if Power[0] == Input1_P == 0:
            scale_Input1_P = 0
        else:
            scale_Input1_P = 'null'

    if Power[1] != 0:
        scale_Input1_Q = abs((Input1_Q - Power[1]) / Power[1])
    else:
        if Power[1] == Input1_Q == 0:
            scale_Input1_Q = 0
        else:
            scale_Input1_Q = 'null'

    if Power[2] != 0:
        scale_Input1_S = abs((Input1_S - Power[2]) / Power[2])
    else:
        if Power[2] == Input1_S == 0:
            scale_Input1_S = 0
        else:
            scale_Input1_S = 'null'

    if Power[3] != 0:
        scale_Input1_PF = abs((Input1_PF - Power[3]) / Power[3])
    else:
        if Power[3] == Input1_PF == 0:
            scale_Input1_PF = 0
        else:
            scale_Input1_PF = 'null'

    Input2_P = read_input_channel_2_active_power(Power[4], times=40)
    Input2_Q = read_input_channel_2_reactive_power(Power[5], times=40)
    Input2_S = read_input_channel_2_apparent_power(Power[6], times=40)
    Input2_PF = read_input_channel_2_power_factor(Power[7], times=40)

    if Power[4] != 0:
        scale_Input2_P = abs((Input2_P - Power[4]) / Power[4])
    else:
        if Power[4] == Input2_P == 0:
            scale_Input2_P = 0
        else:
            scale_Input2_P = 'null'

    if Power[5] != 0:
        scale_Input2_Q = abs((Input2_Q - Power[5]) / Power[5])
    else:
        if Power[5] == Input2_Q == 0:
            scale_Input2_Q = 0
        else:
            scale_Input2_Q = 'null'

    if Power[6] != 0:
        scale_Input2_S = abs((Input2_S - Power[6]) / Power[6])
    else:
        if Power[6] == Input2_S == 0:
            scale_Input2_S = 0
        else:
            scale_Input2_S = 'null'

    if Power[7] != 0:
        scale_Input2_PF = abs((Input2_PF - Power[7]) / Power[7])
    else:
        if Power[7] == Input2_PF == 0:
            scale_Input2_PF = 0
        else:
            scale_Input2_PF = 'null'

    Input3_P = read_input_channel_3_active_power(Power[8], times=40)
    Input3_Q = read_input_channel_3_reactive_power(Power[9], times=40)
    Input3_S = read_input_channel_3_apparent_power(Power[10], times=40)
    Input3_PF = read_input_channel_3_power_factor(Power[11], times=40)

    if Power[8] != 0:
        scale_Input3_P = abs((Input3_P - Power[8]) / Power[8])
    else:
        if Power[8] == Input3_P == 0:
            scale_Input3_P = 0
        else:
            scale_Input3_P = 'null'

    if Power[9] != 0:
        scale_Input3_Q = abs((Input3_Q - Power[9]) / Power[9])
    else:
        if Power[9] == Input3_Q == 0:
            scale_Input3_Q = 0
        else:
            scale_Input3_Q = 'null'

    if Power[10] != 0:
        scale_Input3_S = abs((Input3_S - Power[10]) / Power[10])
    else:
        if Power[10] == Input3_S == 0:
            scale_Input3_S = 0
        else:
            scale_Input3_S = 'null'

    if Power[11] != 0:
        scale_Input3_PF = abs((Input3_PF - Power[11]) / Power[11])
    else:
        if Power[11] == Input3_PF == 0:
            scale_Input3_PF = 0
        else:
            scale_Input3_PF = 'null'

    power_list.extend(
        [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [Input1_P, scale_Input1_P], [Input1_Q, scale_Input1_Q],
         [Input1_S, scale_Input1_S], [Input1_PF, scale_Input1_PF], [Input2_P, scale_Input2_P],
         [Input2_Q, scale_Input2_Q], [Input2_S, scale_Input2_S], [Input2_PF, scale_Input2_PF],
         [Input3_P, scale_Input3_P], [Input3_Q, scale_Input3_Q],
         [Input3_S, scale_Input3_S], [Input3_PF, scale_Input3_PF]])

    if User_Channel == 'user1':
        User1_P = read_user_channel_1_active_power(Power[12], times=40)
        User1_Q = read_user_channel_1_reactive_power(Power[13], times=40)
        User1_S = read_user_channel_1_apparent_power(Power[14], times=40)
        User1_PF = read_user_channel_1_power_factor(Power[15], times=40)

        if Power[12] != 0:
            scale_User1_P = abs((User1_P - Power[12]) / Power[12])
        else:
            if Power[12] == User1_P == 0:
                scale_User1_P = 0
            else:
                scale_User1_P = 'null'

        if Power[13] != 0:
            scale_User1_Q = abs((User1_Q - Power[13]) / Power[13])
        else:
            if Power[13] == User1_Q == 0:
                scale_User1_Q = 0
            else:
                scale_User1_Q = 'null'

        if Power[14] != 0:
            scale_User1_S = abs((User1_S - Power[14]) / Power[14])
        else:
            if Power[14] == User1_S == 0:
                scale_User1_S = 0
            else:
                scale_User1_S = 'null'

        if Power[15] != 0:
            scale_User1_PF = abs((User1_PF - Power[15]) / Power[15])
        else:
            if Power[15] == User1_PF == 0:
                scale_User1_PF = 0
            else:
                scale_User1_PF = 'null'

        power_list.extend(
            [[User1_P, scale_User1_P], [User1_Q, scale_User1_Q], [User1_S, scale_User1_S], [User1_PF, scale_User1_PF]])
    if User_Channel == 'user1,user2,user3':
        User1_P = read_user_channel_1_active_power(Power[0], times=40)
        User1_Q = read_user_channel_1_reactive_power(Power[1], times=40)
        User1_S = read_user_channel_1_apparent_power(Power[2], times=40)
        User1_PF = read_user_channel_1_power_factor(Power[3], times=40)

        if Power[0] != 0:
            scale_User1_P = abs((User1_P - Power[0]) / Power[0])
        else:
            if Power[0] == User1_P == 0:
                scale_User1_P = 0
            else:
                scale_User1_P = 'null'

        if Power[1] != 0:
            scale_User1_Q = abs((User1_Q - Power[1]) / Power[1])
        else:
            if Power[1] == User1_Q == 0:
                scale_User1_Q = 0
            else:
                scale_User1_Q = 'null'

        if Power[2] != 0:
            scale_User1_S = abs((User1_S - Power[2]) / Power[2])
        else:
            if Power[2] == User1_S == 0:
                scale_User1_S = 0
            else:
                scale_User1_S = 'null'

        if Power[3] != 0:
            scale_User1_PF = abs((User1_PF - Power[3]) / Power[3])
        else:
            if Power[3] == User1_PF == 0:
                scale_User1_PF = 0
            else:
                scale_User1_PF = 'null'

        User2_P = read_user_channel_2_active_power(Power[4], times=40)
        User2_Q = read_user_channel_2_reactive_power(Power[5], times=40)
        User2_S = read_user_channel_2_apparent_power(Power[6], times=40)
        User2_PF = read_user_channel_2_power_factor(Power[7], times=40)

        if Power[4] != 0:
            scale_User2_P = abs((User2_P - Power[4]) / Power[4])
        else:
            if Power[4] == User2_P == 0:
                scale_User2_P = 0
            else:
                scale_User2_P = 'null'

        if Power[5] != 0:
            scale_User2_Q = abs((User2_Q - Power[5]) / Power[5])
        else:
            if Power[5] == User2_Q == 0:
                scale_User2_Q = 0
            else:
                scale_User2_Q = 'null'

        if Power[6] != 0:
            scale_User2_S = abs((User2_S - Power[6]) / Power[6])
        else:
            if Power[6] == User2_S == 0:
                scale_User2_S = 0
            else:
                scale_User2_S = 'null'

        if Power[7] != 0:
            scale_User2_PF = abs((User2_PF - Power[7]) / Power[7])
        else:
            if Power[7] == User2_PF == 0:
                scale_User2_PF = 0
            else:
                scale_User2_PF = 'null'

        User3_P = read_user_channel_3_active_power(Power[8], times=40)
        User3_Q = read_user_channel_3_reactive_power(Power[9], times=40)
        User3_S = read_user_channel_3_apparent_power(Power[10], times=40)
        User3_PF = read_user_channel_3_power_factor(Power[11], times=40)

        if Power[8] != 0:
            scale_User3_P = abs((User3_P - Power[8]) / Power[8])
        else:
            if Power[8] == User3_P == 0:
                scale_User3_P = 0
            else:
                scale_User3_P = 'null'

        if Power[9] != 0:
            scale_User3_Q = abs((User3_Q - Power[9]) / Power[9])
        else:
            if Power[9] == User3_Q == 0:
                scale_User3_Q = 0
            else:
                scale_User3_Q = 'null'

        if Power[10] != 0:
            scale_User3_S = abs((User3_S - Power[10]) / Power[10])
        else:
            if Power[10] == User3_S == 0:
                scale_User3_S = 0
            else:
                scale_User3_S = 'null'

        if Power[11] != 0:
            scale_User3_PF = abs((User3_PF - Power[11]) / Power[11])
        else:
            if Power[11] == User3_PF == 0:
                scale_User3_PF = 0
            else:
                scale_User3_PF = 'null'

        power_list.extend(
            [[User1_P, scale_User1_P], [User1_Q, scale_User1_Q], [User1_S, scale_User1_S], [User1_PF, scale_User1_PF],
             [User2_P, scale_User2_P], [User2_Q, scale_User2_Q], [User2_S, scale_User2_S], [User2_PF, scale_User2_PF],
             [User3_P, scale_User3_P], [User3_Q, scale_User3_Q], [User3_S, scale_User3_S], [User3_PF, scale_User3_PF]])
    return power_list


def set_reactive_power_calculation_methodme(value):
    address = basic_setting['Reactive Power Calculation Method']['Start(Dec)']
    count = basic_setting['Reactive Power Calculation Method']['Reg']
    ret = ModbusClient.write_registers(address=address, values=value, slave=1)
    if '(4315,1)' not in str(ret):
        logging.error('set_reactive_power_calculation_methodme fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=address, count=count, slave=1)
    if ret[0] == value:
        return True
    return False


def set_service_configuration(value):
    address = basic_setting['Service Configuration']['Start(Dec)']
    count = basic_setting['Service Configuration']['Reg']
    ret = ModbusClient.write_registers(address=address, values=value, slave=1)
    if '(4162,1)' not in str(ret):
        logging.error('set_service_configuration fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=address, count=count, slave=1)
    if ret[0] == value:
        return True
    return False


def set_clear_energy(value):
    '''
    :param value: 0: None 1: Clearing
    :return:
    '''
    address = basic_setting['Clear energy']['Start(Dec)']
    count = basic_setting['Clear energy']['Reg']
    ret = ModbusClient.write_registers(address=address, values=value, slave=1)
    if '(4400,1)' not in str(ret):
        logging.error('set_clear_energy fail, ret is:{}'.format(ret))
        return False
    return True


def set_device_reboot(value):
    '''
    :param value: 0: None 1: Clearing
    :return:
    '''
    address = basic_setting['Device Reboot']['Start(Dec)']
    count = basic_setting['Device Reboot']['Reg']
    ret = ModbusClient.write_registers(address=address, values=value, slave=1)
    if '(4420,1)' not in str(ret):
        logging.error('set_device_reboot fail, ret is:{}'.format(ret))
        return False
    return True


def set_channle2_voltage_assignment(value):
    '''
    :param value: 1: Vb 2: Vc
    :return:
    '''

    ret = ModbusClient.write_registers(address=4175, values=value, slave=1)
    if '(4175,1)' not in str(ret):
        logging.error('Set_Channle2_voltage_assignment fail, ret is:{}'.format(ret))
        return False
    return True


def set_channle3_voltage_assignment(value):
    '''
    :param value: 1: Vb 2: Vc
    :return:
    '''
    ret = ModbusClient.write_registers(address=4180, values=value, slave=1)
    if '(4180,1)' not in str(ret):
        logging.error('set_channle3_voltage_assignment fail, ret is:{}'.format(ret))
        return False
    return True


def set_phase_order(value):
    '''
    :param value: 0: ABC 1: ACB
    :return:
    '''
    address = basic_setting['Phase Order']['Start(Dec)']
    count = basic_setting['Phase Order']['Reg']
    ret = ModbusClient.write_registers(address=address, values=value, slave=1)
    if '(4316,1)' not in str(ret):
        logging.error('set_phase_order fail, ret is:{}'.format(ret))
        return False
    return True


def read_phase_a_energy():
    address = energy['Phase A active energy import']['Start(Dec)']
    count_reg = energy['Phase A active energy import']['Reg']
    Phase_A_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    if energy_charge == "resp is error":
        energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Phase_A_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_active_energy_import, 16)
    Phase_A_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_active_energy_export, 16)
    Phase_A_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace(
        '0x',
        '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_active_energy_net, 16)
    Phase_A_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_active_energy_total, 16)
    Phase_A_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_reactive_energy_import, 16)
    Phase_A_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_reactive_energy_export, 16)
    Phase_A_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_reactive_energy_net, 16)
    Phase_A_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_reactive_energy_total, 16)
    Phase_A_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_A_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Phase_A_apparent_energy, 16)
    Phase_A_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Phase_A_active_energy_import, Phase_A_active_energy_export , Phase_A_active_energy_net, Phase_A_active_energy_total, Phase_A_reactive_energy_import, Phase_A_reactive_energy_export, Phase_A_reactive_energy_net, Phase_A_reactive_energy_total ret is:{}'.format(
            (Phase_A_active_energy_import, Phase_A_active_energy_export, Phase_A_active_energy_net,
             Phase_A_active_energy_total, Phase_A_reactive_energy_import, Phase_A_reactive_energy_export,
             Phase_A_reactive_energy_net,
             Phase_A_reactive_energy_total, Phase_A_apparent_energy)))
    Phase_A_Energy_list.extend(
        [Phase_A_active_energy_import, Phase_A_active_energy_export, Phase_A_active_energy_net,
         Phase_A_active_energy_total,
         Phase_A_reactive_energy_import, Phase_A_reactive_energy_export,
         Phase_A_reactive_energy_net,
         Phase_A_reactive_energy_total, Phase_A_apparent_energy])
    return Phase_A_Energy_list


def read_phase_b_energy():
    address = energy['Phase B active energy import']['Start(Dec)']
    count_reg = energy['Phase B active energy import']['Reg']
    Phase_B_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Phase_B_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_active_energy_import, 16)
    Phase_B_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_active_energy_export, 16)
    Phase_B_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace('0x',
                                                                                                                 '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_active_energy_net, 16)
    Phase_B_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_active_energy_total, 16)
    Phase_B_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_reactive_energy_import, 16)
    Phase_B_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_reactive_energy_export, 16)
    Phase_B_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_reactive_energy_net, 16)
    Phase_B_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_reactive_energy_total, 16)
    Phase_B_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_B_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Phase_B_apparent_energy, 16)
    Phase_B_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Phase_B_active_energy_import, Phase_B_active_energy_export , Phase_B_active_energy_net, Phase_B_active_energy_total, Phase_B_reactive_energy_import, Phase_B_reactive_energy_export, Phase_B_reactive_energy_net, Phase_B_reactive_energy_total ret is:{}'.format(
            (Phase_B_active_energy_import, Phase_B_active_energy_export, Phase_B_active_energy_net,
             Phase_B_active_energy_total, Phase_B_reactive_energy_import, Phase_B_reactive_energy_export,
             Phase_B_reactive_energy_net,
             Phase_B_reactive_energy_total, Phase_B_apparent_energy)))
    Phase_B_Energy_list.extend([Phase_B_active_energy_import, Phase_B_active_energy_export, Phase_B_active_energy_net,
                                Phase_B_active_energy_total,
                                Phase_B_reactive_energy_import, Phase_B_reactive_energy_export,
                                Phase_B_reactive_energy_net,
                                Phase_B_reactive_energy_total, Phase_B_apparent_energy])
    return Phase_B_Energy_list


def read_phase_c_energy():
    address = energy['Phase C active energy import']['Start(Dec)']
    count_reg = energy['Phase C active energy import']['Reg']
    Phase_C_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Phase_C_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_active_energy_import, 16)
    Phase_C_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_active_energy_export, 16)
    Phase_C_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace('0x',
                                                                                                                 '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_active_energy_net, 16)
    Phase_C_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_active_energy_total, 16)
    Phase_C_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_reactive_energy_import, 16)
    Phase_C_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_reactive_energy_export, 16)
    Phase_C_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_reactive_energy_net, 16)
    Phase_C_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_reactive_energy_total, 16)
    Phase_C_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Phase_C_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Phase_C_apparent_energy, 16)
    Phase_C_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Phase_C_active_energy_import, Phase_C_active_energy_export , Phase_C_active_energy_net, Phase_C_active_energy_total, Phase_C_reactive_energy_import, Phase_C_reactive_energy_export, Phase_C_reactive_energy_net, Phase_C_reactive_energy_total ret is:{}'.format(
            (Phase_C_active_energy_import, Phase_C_active_energy_export, Phase_C_active_energy_net,
             Phase_C_active_energy_total, Phase_C_reactive_energy_import, Phase_C_reactive_energy_export,
             Phase_C_reactive_energy_net,
             Phase_C_reactive_energy_total, Phase_C_apparent_energy)))
    Phase_C_Energy_list.extend([Phase_C_active_energy_import, Phase_C_active_energy_export, Phase_C_active_energy_net,
                                Phase_C_active_energy_total,
                                Phase_C_reactive_energy_import, Phase_C_reactive_energy_export,
                                Phase_C_reactive_energy_net,
                                Phase_C_reactive_energy_total, Phase_C_apparent_energy])
    return Phase_C_Energy_list


def read_system_energy():
    address = energy['System active energy import']['Start(Dec)']
    count_reg = energy['System active energy import']['Reg']
    System_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    System_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace('0x',
                                                                                                                   '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(System_active_energy_import, 16)
    System_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace('0x',
                                                                                                                   '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(System_active_energy_export, 16)
    System_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace('0x',
                                                                                                                '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(System_active_energy_net, 16)
    System_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(System_active_energy_total, 16)
    System_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(System_reactive_energy_import, 16)
    System_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(System_reactive_energy_export, 16)
    System_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(System_reactive_energy_net, 16)
    System_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(System_reactive_energy_total, 16)
    System_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    System_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(System_apparent_energy, 16)
    System_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'System_active_energy_import, System_active_energy_export , System_active_energy_net, System_active_energy_total, System_reactive_energy_import, System_reactive_energy_export, System_reactive_energy_net, System_reactive_energy_total ret is:{}'.format(
            (System_active_energy_import, System_active_energy_export, System_active_energy_net,
             System_active_energy_total, System_reactive_energy_import, System_reactive_energy_export,
             System_reactive_energy_net,
             System_reactive_energy_total, System_apparent_energy)))
    System_Energy.extend(
        [System_active_energy_import, System_active_energy_export, System_active_energy_net, System_active_energy_total,
         System_reactive_energy_import, System_reactive_energy_export, System_reactive_energy_net,
         System_reactive_energy_total, System_apparent_energy])
    return System_Energy


def read_input_channel_1_energy():
    address = energy['Input Channel 1 active energy import']['Start(Dec)']
    count_reg = energy['Input Channel 1 active energy import']['Reg']
    Input_Channel_1_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Input_Channel_1_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_active_energy_import, 16)
    Input_Channel_1_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_active_energy_export, 16)
    Input_Channel_1_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(
        energy_charge[9]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_active_energy_net, 16)
    Input_Channel_1_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_active_energy_total, 16)
    Input_Channel_1_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_reactive_energy_import, 16)
    Input_Channel_1_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_reactive_energy_export, 16)
    Input_Channel_1_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_reactive_energy_net, 16)
    Input_Channel_1_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_reactive_energy_total, 16)
    Input_Channel_1_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_1_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(
        energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_1_apparent_energy, 16)
    Input_Channel_1_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Input_Channel_1_active_energy_import, Input_Channel_1_active_energy_export , Input_Channel_1_active_energy_net, Input_Channel_1_active_energy_total, Input_Channel_1_reactive_energy_import, Input_Channel_1_reactive_energy_export, Input_Channel_1_reactive_energy_net, Input_Channel_1_reactive_energy_total ret is:{}'.format(
            (Input_Channel_1_active_energy_import, Input_Channel_1_active_energy_export,
             Input_Channel_1_active_energy_net,
             Input_Channel_1_active_energy_total, Input_Channel_1_reactive_energy_import,
             Input_Channel_1_reactive_energy_export,
             Input_Channel_1_reactive_energy_net,
             Input_Channel_1_reactive_energy_total, Input_Channel_1_apparent_energy)))
    Input_Channel_1_Energy.extend(
        [Input_Channel_1_active_energy_import, Input_Channel_1_active_energy_export, Input_Channel_1_active_energy_net,
         Input_Channel_1_active_energy_total,
         Input_Channel_1_reactive_energy_import, Input_Channel_1_reactive_energy_export,
         Input_Channel_1_reactive_energy_net,
         Input_Channel_1_reactive_energy_total, Input_Channel_1_apparent_energy])
    return Input_Channel_1_Energy


def read_input_channel_2_energy():
    address = energy['Input Channel 1 active energy import']['Start(Dec)'] + 36
    count_reg = energy['Input Channel 1 active energy import']['Reg']
    Input_Channel_2_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Input_Channel_2_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_active_energy_import, 16)
    Input_Channel_2_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_active_energy_export, 16)
    Input_Channel_2_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(
        energy_charge[9]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_active_energy_net, 16)
    Input_Channel_2_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_active_energy_total, 16)
    Input_Channel_2_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_reactive_energy_import, 16)
    Input_Channel_2_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_reactive_energy_export, 16)
    Input_Channel_2_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_reactive_energy_net, 16)
    Input_Channel_2_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_reactive_energy_total, 16)
    Input_Channel_2_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_2_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(
        energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_2_apparent_energy, 16)
    Input_Channel_2_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Input_Channel_2_active_energy_import, Input_Channel_2_active_energy_export , Input_Channel_2_active_energy_net, Input_Channel_2_active_energy_total, Input_Channel_2_reactive_energy_import, Input_Channel_2_reactive_energy_export, Input_Channel_2_reactive_energy_net, Input_Channel_2_reactive_energy_total ret is:{}'.format(
            (Input_Channel_2_active_energy_import, Input_Channel_2_active_energy_export,
             Input_Channel_2_active_energy_net,
             Input_Channel_2_active_energy_total, Input_Channel_2_reactive_energy_import,
             Input_Channel_2_reactive_energy_export,
             Input_Channel_2_reactive_energy_net,
             Input_Channel_2_reactive_energy_total, Input_Channel_2_apparent_energy)))
    Input_Channel_2_Energy.extend(
        [Input_Channel_2_active_energy_import, Input_Channel_2_active_energy_export, Input_Channel_2_active_energy_net,
         Input_Channel_2_active_energy_total,
         Input_Channel_2_reactive_energy_import, Input_Channel_2_reactive_energy_export,
         Input_Channel_2_reactive_energy_net,
         Input_Channel_2_reactive_energy_total, Input_Channel_2_apparent_energy])
    return Input_Channel_2_Energy


def read_input_channel_3_energy():
    address = energy['Input Channel 1 active energy import']['Start(Dec)'] + 72
    count_reg = energy['Input Channel 1 active energy import']['Reg']
    Input_Channel_3_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    Input_Channel_3_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_active_energy_import, 16)
    Input_Channel_3_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_active_energy_export, 16)
    Input_Channel_3_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(
        energy_charge[9]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_active_energy_net, 16)
    Input_Channel_3_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_active_energy_total, 16)
    Input_Channel_3_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_reactive_energy_import, 16)
    Input_Channel_3_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_reactive_energy_export, 16)
    Input_Channel_3_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_reactive_energy_net, 16)
    Input_Channel_3_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_reactive_energy_total, 16)
    Input_Channel_3_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    Input_Channel_3_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(
        energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(Input_Channel_3_apparent_energy, 16)
    Input_Channel_3_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'Input_Channel_3_active_energy_import, Input_Channel_3_active_energy_export , Input_Channel_3_active_energy_net, Input_Channel_3_active_energy_total, Input_Channel_3_reactive_energy_import, Input_Channel_3_reactive_energy_export, Input_Channel_3_reactive_energy_net, Input_Channel_3_reactive_energy_total ret is:{}'.format(
            (Input_Channel_3_active_energy_import, Input_Channel_3_active_energy_export,
             Input_Channel_3_active_energy_net,
             Input_Channel_3_active_energy_total, Input_Channel_3_reactive_energy_import,
             Input_Channel_3_reactive_energy_export,
             Input_Channel_3_reactive_energy_net,
             Input_Channel_3_reactive_energy_total, Input_Channel_3_apparent_energy)))
    Input_Channel_3_Energy.extend(
        [Input_Channel_3_active_energy_import, Input_Channel_3_active_energy_export, Input_Channel_3_active_energy_net,
         Input_Channel_3_active_energy_total,
         Input_Channel_3_reactive_energy_import, Input_Channel_3_reactive_energy_export,
         Input_Channel_3_reactive_energy_net,
         Input_Channel_3_reactive_energy_total, Input_Channel_3_apparent_energy])
    return Input_Channel_3_Energy


def read_user_channel_1_energy():
    address = energy['User Channel 1 active energy import']['Start(Dec)']
    count_reg = energy['User Channel 1 active energy import']['Reg']
    User_Channel_1_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    User_Channel_1_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_active_energy_import, 16)
    User_Channel_1_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_active_energy_export, 16)
    User_Channel_1_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace(
        '0x',
        '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_active_energy_net, 16)
    User_Channel_1_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_active_energy_total, 16)
    User_Channel_1_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_reactive_energy_import, 16)
    User_Channel_1_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_reactive_energy_export, 16)
    User_Channel_1_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_reactive_energy_net, 16)
    User_Channel_1_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_reactive_energy_total, 16)
    User_Channel_1_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_1_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_1_apparent_energy, 16)
    User_Channel_1_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'User_Channel_1_active_energy_import, User_Channel_1_active_energy_export , User_Channel_1_active_energy_net, User_Channel_1_active_energy_total, User_Channel_1_reactive_energy_import, User_Channel_1_reactive_energy_export, User_Channel_1_reactive_energy_net, User_Channel_1_reactive_energy_total ret is:{}'.format(
            (User_Channel_1_active_energy_import, User_Channel_1_active_energy_export, User_Channel_1_active_energy_net,
             User_Channel_1_active_energy_total, User_Channel_1_reactive_energy_import,
             User_Channel_1_reactive_energy_export,
             User_Channel_1_reactive_energy_net,
             User_Channel_1_reactive_energy_total, User_Channel_1_apparent_energy)))
    User_Channel_1_Energy.extend(
        [User_Channel_1_active_energy_import, User_Channel_1_active_energy_export, User_Channel_1_active_energy_net,
         User_Channel_1_active_energy_total,
         User_Channel_1_reactive_energy_import, User_Channel_1_reactive_energy_export,
         User_Channel_1_reactive_energy_net,
         User_Channel_1_reactive_energy_total, User_Channel_1_apparent_energy])
    return User_Channel_1_Energy


def read_user_channel_2_energy():
    address = energy['User Channel 1 active energy import']['Start(Dec)'] + 36
    count_reg = energy['User Channel 1 active energy import']['Reg']
    User_Channel_2_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    User_Channel_2_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_active_energy_import, 16)
    User_Channel_2_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_active_energy_export, 16)
    User_Channel_2_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace(
        '0x',
        '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_active_energy_net, 16)
    User_Channel_2_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_active_energy_total, 16)
    User_Channel_2_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_reactive_energy_import, 16)
    User_Channel_2_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_reactive_energy_export, 16)
    User_Channel_2_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_reactive_energy_net, 16)
    User_Channel_2_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_reactive_energy_total, 16)
    User_Channel_2_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_2_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_2_apparent_energy, 16)
    User_Channel_2_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'User_Channel_2_active_energy_import, User_Channel_2_active_energy_export , User_Channel_2_active_energy_net, User_Channel_2_active_energy_total, User_Channel_2_reactive_energy_import, User_Channel_2_reactive_energy_export, User_Channel_2_reactive_energy_net, User_Channel_2_reactive_energy_total ret is:{}'.format(
            (User_Channel_2_active_energy_import, User_Channel_2_active_energy_export, User_Channel_2_active_energy_net,
             User_Channel_2_active_energy_total, User_Channel_2_reactive_energy_import,
             User_Channel_2_reactive_energy_export,
             User_Channel_2_reactive_energy_net,
             User_Channel_2_reactive_energy_total, User_Channel_2_apparent_energy)))
    User_Channel_2_Energy.extend(
        [User_Channel_2_active_energy_import, User_Channel_2_active_energy_export, User_Channel_2_active_energy_net,
         User_Channel_2_active_energy_total,
         User_Channel_2_reactive_energy_import, User_Channel_2_reactive_energy_export,
         User_Channel_2_reactive_energy_net,
         User_Channel_2_reactive_energy_total, User_Channel_2_apparent_energy])
    return User_Channel_2_Energy


def read_user_channel_3_energy():
    address = energy['User Channel 1 active energy import']['Start(Dec)'] + 72
    count_reg = energy['User Channel 1 active energy import']['Reg']
    User_Channel_3_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=address, count=36, slave=1)
    User_Channel_3_active_energy_import = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(
        energy_charge[1]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_active_energy_import, 16)
    User_Channel_3_active_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_active_energy_export = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(
        energy_charge[5]).replace('0x',
                                  '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_active_energy_export, 16)
    User_Channel_3_active_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_active_energy_net = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace(
        '0x',
        '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_active_energy_net, 16)
    User_Channel_3_active_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_active_energy_total = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(
        energy_charge[13]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_active_energy_total, 16)
    User_Channel_3_active_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_reactive_energy_import = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(
        energy_charge[17]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_reactive_energy_import, 16)
    User_Channel_3_reactive_energy_import = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_reactive_energy_export = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(
        energy_charge[21]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_reactive_energy_export, 16)
    User_Channel_3_reactive_energy_export = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_reactive_energy_net = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(
        energy_charge[25]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_reactive_energy_net, 16)
    User_Channel_3_reactive_energy_net = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_reactive_energy_total = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(
        energy_charge[29]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_reactive_energy_total, 16)
    User_Channel_3_reactive_energy_total = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    User_Channel_3_apparent_energy = hex(energy_charge[32]).replace('0x', '').zfill(4) + hex(energy_charge[33]).replace(
        '0x', '').zfill(
        4) + hex(energy_charge[34]).replace('0x', '').zfill(4) + hex(energy_charge[35]).replace('0x', '').zfill(4)
    integer_num = int(User_Channel_3_apparent_energy, 16)
    User_Channel_3_apparent_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'User_Channel_3_active_energy_import, User_Channel_3_active_energy_export , User_Channel_3_active_energy_net, User_Channel_3_active_energy_total, User_Channel_3_reactive_energy_import, User_Channel_3_reactive_energy_export, User_Channel_3_reactive_energy_net, User_Channel_3_reactive_energy_total ret is:{}'.format(
            (User_Channel_3_active_energy_import, User_Channel_3_active_energy_export, User_Channel_3_active_energy_net,
             User_Channel_3_active_energy_total, User_Channel_3_reactive_energy_import,
             User_Channel_3_reactive_energy_export,
             User_Channel_3_reactive_energy_net,
             User_Channel_3_reactive_energy_total, User_Channel_3_apparent_energy)))
    User_Channel_3_Energy.extend(
        [User_Channel_3_active_energy_import, User_Channel_3_active_energy_export, User_Channel_3_active_energy_net,
         User_Channel_3_active_energy_total,
         User_Channel_3_reactive_energy_import, User_Channel_3_reactive_energy_export,
         User_Channel_3_reactive_energy_net,
         User_Channel_3_reactive_energy_total, User_Channel_3_apparent_energy])
    return User_Channel_3_Energy


def energy_standard_value_calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time):
    '''

    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param Time:
    :return: Phase_A_P_E_import, Phase_A_P_E_export, Phase_A_P_E_net, Phase_A_P_E_total, Phase_A_Q_E_import,
         Phase_A_Q_E_export, Phase_A_Q_E_net, Phase_A_Q_E_total, Phase_A_S_E, Phase_B_P_E_import, Phase_B_P_E_export,
         Phase_B_P_E_net, Phase_B_P_E_total, Phase_B_Q_E_import, Phase_B_Q_E_export, Phase_B_Q_E_net, Phase_B_Q_E_total,
         Phase_B_S_E, Phase_C_P_E_import, Phase_C_P_E_export, Phase_C_P_E_net, Phase_C_P_E_total, Phase_C_Q_E_import,
         Phase_C_Q_E_export, Phase_C_Q_E_net, Phase_C_Q_E_total, Phase_C_S_E, System_P_E_import, System_P_E_export,
         System_P_E_net, System_P_E_total, System_Q_E_import, System_Q_E_export, System_Q_E_net, System_Q_E_total,
         System_S_E
    '''
    Energy_Standard_Value_list = []
    Power_standard_value_list = power_standard_value_calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang,
                                                               Ib_ang, Ic_ang)
    Phase_A_P = Power_standard_value_list[0]
    Phase_A_Q = Power_standard_value_list[1]
    Phase_A_S = Power_standard_value_list[2]
    if Phase_A_P > 0:
        Phase_A_P_E_import = abs(Phase_A_P * (Time / 60))
        Phase_A_P_E_export = 0
    else:
        Phase_A_P_E_import = 0
        Phase_A_P_E_export = abs(Phase_A_P * (Time / 60))
    Phase_A_P_E_net = Phase_A_P_E_import - Phase_A_P_E_export
    Phase_A_P_E_total = Phase_A_P_E_import + Phase_A_P_E_export

    if Phase_A_Q > 0:
        Phase_A_Q_E_import = abs(Phase_A_Q * (Time / 60))
        Phase_A_Q_E_export = 0
    else:
        Phase_A_Q_E_import = 0
        Phase_A_Q_E_export = abs(Phase_A_Q * (Time / 60))
    Phase_A_Q_E_net = Phase_A_Q_E_import - Phase_A_Q_E_export
    Phase_A_Q_E_total = Phase_A_Q_E_import + Phase_A_Q_E_export

    Phase_A_S_E = Phase_A_S * (Time / 60)

    Phase_B_P = Power_standard_value_list[4]
    Phase_B_Q = Power_standard_value_list[5]
    Phase_B_S = Power_standard_value_list[6]
    if Phase_B_P > 0:
        Phase_B_P_E_import = abs(Phase_B_P * (Time / 60))
        Phase_B_P_E_export = 0
    else:
        Phase_B_P_E_import = 0
        Phase_B_P_E_export = abs(Phase_B_P * (Time / 60))
    Phase_B_P_E_net = Phase_B_P_E_import - Phase_B_P_E_export
    Phase_B_P_E_total = Phase_B_P_E_import + Phase_B_P_E_export

    if Phase_B_Q > 0:
        Phase_B_Q_E_import = abs(Phase_B_Q * (Time / 60))
        Phase_B_Q_E_export = 0
    else:
        Phase_B_Q_E_import = 0
        Phase_B_Q_E_export = abs(Phase_B_Q * (Time / 60))
    Phase_B_Q_E_net = Phase_B_Q_E_import - Phase_B_Q_E_export
    Phase_B_Q_E_total = Phase_B_Q_E_import + Phase_B_Q_E_export

    Phase_B_S_E = Phase_B_S * (Time / 60)

    Phase_C_P = Power_standard_value_list[8]
    Phase_C_Q = Power_standard_value_list[9]
    Phase_C_S = Power_standard_value_list[10]
    if Phase_C_P > 0:
        Phase_C_P_E_import = abs(Phase_C_P * (Time / 60))
        Phase_C_P_E_export = 0
    else:
        Phase_C_P_E_import = 0
        Phase_C_P_E_export = abs(Phase_C_P * (Time / 60))
    Phase_C_P_E_net = Phase_C_P_E_import - Phase_C_P_E_export
    Phase_C_P_E_total = Phase_C_P_E_import + Phase_C_P_E_export

    if Phase_C_Q > 0:
        Phase_C_Q_E_import = abs(Phase_C_Q * (Time / 60))
        Phase_C_Q_E_export = 0
    else:
        Phase_C_Q_E_import = 0
        Phase_C_Q_E_export = abs(Phase_C_Q * (Time / 60))
    Phase_C_Q_E_net = Phase_C_Q_E_import - Phase_C_Q_E_export
    Phase_C_Q_E_total = Phase_C_Q_E_import + Phase_C_Q_E_export

    Phase_C_S_E = Phase_C_S * (Time / 60)

    System_P_E_import = Phase_A_P_E_import + Phase_B_P_E_import + Phase_C_P_E_import
    System_P_E_export = Phase_A_P_E_export + Phase_B_P_E_export + Phase_C_P_E_export
    System_P_E_net = Phase_A_P_E_net + Phase_B_P_E_net + Phase_C_P_E_net
    System_P_E_total = Phase_A_P_E_total + Phase_B_P_E_total + Phase_C_P_E_total

    System_Q_E_import = Phase_A_Q_E_import + Phase_B_Q_E_import + Phase_C_Q_E_import
    System_Q_E_export = Phase_A_Q_E_export + Phase_B_Q_E_export + Phase_C_Q_E_export
    System_Q_E_net = Phase_A_Q_E_net + Phase_B_Q_E_net + Phase_C_Q_E_net
    System_Q_E_total = Phase_A_Q_E_total + Phase_B_Q_E_total + Phase_C_Q_E_total

    System_S_E = Phase_A_S_E + Phase_B_S_E + Phase_C_S_E

    Energy_Standard_Value_list.extend(
        [Phase_A_P_E_import, Phase_A_P_E_export, Phase_A_P_E_net, Phase_A_P_E_total, Phase_A_Q_E_import,
         Phase_A_Q_E_export, Phase_A_Q_E_net, Phase_A_Q_E_total, Phase_A_S_E, Phase_B_P_E_import, Phase_B_P_E_export,
         Phase_B_P_E_net, Phase_B_P_E_total, Phase_B_Q_E_import, Phase_B_Q_E_export, Phase_B_Q_E_net, Phase_B_Q_E_total,
         Phase_B_S_E, Phase_C_P_E_import, Phase_C_P_E_export, Phase_C_P_E_net, Phase_C_P_E_total, Phase_C_Q_E_import,
         Phase_C_Q_E_export, Phase_C_Q_E_net, Phase_C_Q_E_total, Phase_C_S_E, System_P_E_import, System_P_E_export,
         System_P_E_net, System_P_E_total, System_Q_E_import, System_Q_E_export, System_Q_E_net, System_Q_E_total,
         System_S_E])
    return Energy_Standard_Value_list


def read_energy_scale(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time, Service):
    '''

    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param Time:
    :param Service:
    :return:
    :Read_Energy_list:4100_Energy_list
    :Energy_scale_list
    '''
    Energy_Standard_Value_list = energy_standard_value_calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang,
                                                                 Ib_ang, Ic_ang, Time)
    Phase_A_Energy_list = read_phase_a_energy()
    Phase_B_Energy_list = read_phase_b_energy()
    Phase_C_Energy_list = read_phase_c_energy()
    System_Energy_list = read_system_energy()
    Input_Channel_1_Energy = read_input_channel_1_energy()
    Input_Channel_2_Energy = read_input_channel_2_energy()
    Input_Channel_3_Energy = read_input_channel_3_energy()
    User_Channel_1_Energy = read_user_channel_1_energy()
    User_Channel_2_Energy = read_user_channel_2_energy()
    User_Channel_3_Energy = read_user_channel_3_energy()
    Read_Energy_list = []
    Read_Energy_list.extend(
        Phase_A_Energy_list + Phase_B_Energy_list + Phase_C_Energy_list + System_Energy_list + Input_Channel_1_Energy +
        Input_Channel_2_Energy + Input_Channel_3_Energy + User_Channel_1_Energy + User_Channel_2_Energy + User_Channel_3_Energy)
    Energy_scale_list = []
    for i in range(len(Energy_Standard_Value_list)):
        if Energy_Standard_Value_list[i] != 0:
            Energy_scale = abs((Read_Energy_list[i] - Energy_Standard_Value_list[i]) / Energy_Standard_Value_list[i])
            Energy_scale_list.append(Energy_scale)
        elif Energy_Standard_Value_list[i] == Read_Energy_list[i] == 0:
            Energy_scale = 0
            Energy_scale_list.append(Energy_scale)
        else:
            Energy_scale = 'null'
            Energy_scale_list.append(Energy_scale)
    for i in range(36, 63):
        if Energy_Standard_Value_list[i - 36] != 0:
            Energy_scale = abs(
                (Read_Energy_list[i] - Energy_Standard_Value_list[i - 36]) / Energy_Standard_Value_list[i - 36])
            Energy_scale_list.append(Energy_scale)
        elif Energy_Standard_Value_list[i - 36] == Read_Energy_list[i] == 0:
            Energy_scale = 0
            Energy_scale_list.append(Energy_scale)
        else:
            Energy_scale = 'null'
            Energy_scale_list.append(Energy_scale)
    if Service == '3E3p4w':
        for i in range(63, 72):
            if Energy_Standard_Value_list[i - 36] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i - 36]) / Energy_Standard_Value_list[i - 36])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i - 36] == Read_Energy_list[i] == 0:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 'null'
                Energy_scale_list.append(Energy_scale)
        for i in range(72, 90):
            if Read_Energy_list[i] == 0:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 'null'
                Energy_scale_list.append(Energy_scale)
    if Service == '1E1p2w':
        for i in range(63, 90):
            if Energy_Standard_Value_list[i - 63] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i - 63]) / Energy_Standard_Value_list[i - 63])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i - 63] == Read_Energy_list[i] == 0:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 'null'
                Energy_scale_list.append(Energy_scale)
    return Read_Energy_list, Energy_scale_list


def calculate_angle(sequence_component_complex, sequence_component):
    if sequence_component > 0.00001:
        phase = cmath.phase(sequence_component_complex)
        angle = math.degrees(phase)

        if abs(angle) < 0.000001:
            angle = 0

        # 如果角度为负数，调整为 0 到 360 度之间
        if angle < 0:
            angle += 360

        return angle
    else:
        return 0


def sequence_component_calculation(a: float, b: float, c: float, a_angle: float, b_angle: float, c_angle: float):
    ret = []
    va_complex = complex(math.cos(a_angle * math.pi / 180) * a, math.sin(a_angle * math.pi / 180) * a)
    vb_complex = complex(math.cos(b_angle * math.pi / 180) * b, math.sin(b_angle * math.pi / 180) * b)
    vc_complex = complex(math.cos(c_angle * math.pi / 180) * c, math.sin(c_angle * math.pi / 180) * c)
    Rotation_Factor = complex(-1 / 2, math.sqrt(3) / 2)
    Rotation_Factor_square = Rotation_Factor ** 2
    zero_sequence_component_complex = (va_complex + vb_complex + vc_complex) / 3
    zero_sequence_component = abs(zero_sequence_component_complex)
    zero_seq_calculate_angle = round(calculate_angle(zero_sequence_component_complex, zero_sequence_component), 3)

    positive_sequence_component_complex = (va_complex / 3) + ((vb_complex * Rotation_Factor) / 3) + (
            (Rotation_Factor_square * vc_complex) / 3)
    positive_sequence_component = abs(positive_sequence_component_complex)
    positive_seq_calculate_angle = round(
        calculate_angle(positive_sequence_component_complex, positive_sequence_component), 3)

    negative_sequence_component_complex = (va_complex / 3) + ((vb_complex * Rotation_Factor_square) / 3) + (
            (Rotation_Factor * vc_complex) / 3)
    negative_sequence_component = abs(negative_sequence_component_complex)
    negative_seq_calculate_angle = round(
        calculate_angle(negative_sequence_component_complex, negative_sequence_component), 3)
    try:
        VUF_CUF = (round((negative_sequence_component / positive_sequence_component), 10)) * 100
    except:
        VUF_CUF = 0
    if VUF_CUF > 150:
        VUF_CUF = 150

    return round(zero_sequence_component, 3), zero_seq_calculate_angle, round(positive_sequence_component,
                                                                              3), positive_seq_calculate_angle, round(
        negative_sequence_component, 3), negative_seq_calculate_angle, VUF_CUF


def read_voltage_positive_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Positive Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['Voltage Positive Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Positive Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_zero_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Zero Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['Voltage Zero Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Zero Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_negative_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Negative Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['Voltage Negative Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Negative Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_positive_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Positive Sequence Angle']['Start(Dec)']
    count = real_time_addr['Voltage Positive Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Positive Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_zero_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Zero Sequence Angle']['Start(Dec)']
    count = real_time_addr['Voltage Zero Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Zero Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_negative_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Negative Sequence Angle']['Start(Dec)']
    count = real_time_addr['Voltage Negative Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Negative Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_voltage_unbalance_factor_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Voltage Unbalance Factor Magnitude']['Start(Dec)']
    count = real_time_addr['Voltage Unbalance Factor Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Voltage Unbalance Factor Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_positive_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Positive Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Positive Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Positive Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_zero_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Zero Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Zero Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Zero Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_negative_sequence_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Negative Sequence Magnitude']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Negative Sequence Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Negative Sequence Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_positive_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Positive Sequence Angle']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Positive Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Positive Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_zero_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Zero Sequence Angle']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Zero Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Zero Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_negative_sequence_angle(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Negative Sequence Angle']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Negative Sequence Angle']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Negative Sequence Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_user_channel_1_current_unbalance_factor_magnitude(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current Unbalance Factor Magnitude']['Start(Dec)']
    count = real_time_addr['User Channel 1 Current Unbalance Factor Magnitude']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('User Channel 1 Current Unbalance Factor Magnitude ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_4_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 42
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 4 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_5_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 56
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 5 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_6_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 70
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 6 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_7_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 84
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 7 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_8_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 98
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 8 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_9_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + 112
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 9 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_10_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (10 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 10 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_11_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (11 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 11 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_12_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (12 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 12 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_13_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (13 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 13 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_14_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (14 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 14 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_15_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (15 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 15 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_16_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (16 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 16 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_17_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (17 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 17 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_18_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (18 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 18 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_19_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (19 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 19 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_20_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (20 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 20 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_21_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (21 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 21 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_22_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (22 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 22 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_23_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (23 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 23 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_channel_24_current(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    input_channel_add = (24 - 1) * 14
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)'] + input_channel_add
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Input Channel 24 Current ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_phase_sys_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Phase A Current']['Start(Dec)']
    count = 48
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=48, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
        value = [value[i] for i in range(len(value)) if
                 i not in {0, 1, 8, 9, 12, 13, 20, 21, 24, 25, 32, 33, 36, 37, 44, 45}]
        read_list = []
        for i in range(16):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(16):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list


def read_phase_input_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)']
    count = 42
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=42, slave=1)
        logging.info('Input Channel 1 Current ret is:{}'.format(value))
        value = [value[i] for i in range(len(value)) if
                 i not in {0, 1, 8, 9, 12, 13, 14, 15, 22, 23, 26, 27, 28, 29, 36, 37, 40, 41}]
        read_list = []
        for i in range(12):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(12):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list


def read_user_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Current']['Start(Dec)']
    count = 36
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8600, count=36, slave=1)
        logging.info('User Channel 1 Current ret is:{}'.format(value))
        value = [value[i] for i in range(len(value)) if
                 i not in {0, 1, 8, 9, 12, 13, 20, 21, 24, 25, 32, 33}]
        read_list = []
        for i in range(12):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(12):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list


def power_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang,
                                       User_Channel):
    """
    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :return:
    Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P, Phase_C_Q,Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF
    """
    if User_Channel == "3E3p4w":
        Power = []
        Phase_A_P = round(active_power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        Phase_A_Q = round(reactive_power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        Phase_A_S = round(apparent_power_calculate(Va, Ia), 10)
        Phase_A_PF = round(power_factor_calculate(Va_ang, Ia_ang), 10)

        Phase_B_P = round(active_power_calculate(Vb, Ib, Vb_ang, Ib_ang), 10)
        Phase_B_Q = round(reactive_power_calculate(Vb, Ib, Vb_ang, Ib_ang), 10)
        Phase_B_S = round(apparent_power_calculate(Vb, Ib), 10)
        Phase_B_PF = round(power_factor_calculate(Vb_ang, Ib_ang), 10)

        Phase_C_P = round(active_power_calculate(Vc, Ic, Vc_ang, Ic_ang), 10)
        Phase_C_Q = round(reactive_power_calculate(Vc, Ic, Vc_ang, Ic_ang), 10)
        Phase_C_S = round(apparent_power_calculate(Vc, Ic), 10)
        Phase_C_PF = round(power_factor_calculate(Vc_ang, Ic_ang), 10)

        Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
        Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
        Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

        if Sys_S != 0:
            Sys_PF = round(Sys_P / Sys_S, 3)
        else:
            Sys_PF = 0
        Power.extend(
            [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
             Phase_C_Q,
             Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
        return Power
    if User_Channel == "1E1p2w":
        Power = []
        input_1_P = round(active_power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        input_1_Q = round(reactive_power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        input_1_S = round(apparent_power_calculate(Va, Ia), 10)
        input_1_PF = round(power_factor_calculate(Va_ang, Ia_ang), 10)

        input_2_P = round(active_power_calculate(Va, Ib, Va_ang, Ib_ang), 10)
        input_2_Q = round(reactive_power_calculate(Va, Ib, Va_ang, Ib_ang), 10)
        input_2_S = round(apparent_power_calculate(Va, Ib), 10)
        input_2_PF = round(power_factor_calculate(Va_ang, Ib_ang), 10)

        input_3_P = round(active_power_calculate(Va, Ic, Va_ang, Ic_ang), 10)
        input_3_Q = round(reactive_power_calculate(Va, Ic, Va_ang, Ic_ang), 10)
        input_3_S = round(apparent_power_calculate(Va, Ic), 10)
        input_3_PF = round(power_factor_calculate(Va_ang, Ic_ang), 10)

        Phase_A_P = input_1_P + input_2_P + input_3_P
        Phase_A_Q = input_1_Q + input_2_Q + input_3_Q
        Phase_A_S = input_1_S + input_2_S + input_3_S
        if Phase_A_S != 0:
            Phase_A_PF = round(Phase_A_P / Phase_A_S, 3)
        else:
            Phase_A_PF = 0

        Phase_B_P = 0
        Phase_B_Q = 0
        Phase_B_S = 0
        Phase_B_PF = 0

        Phase_C_P = 0
        Phase_C_Q = 0
        Phase_C_S = 0
        Phase_C_PF = 0

        Sys_P = Phase_A_P
        Sys_Q = Phase_A_Q
        Sys_S = Phase_A_S
        Sys_PF = Phase_A_PF
        Power.extend(
            [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
             Phase_C_Q, Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF, input_1_P, input_1_Q, input_1_S,
             input_1_PF, input_2_P, input_2_Q, input_2_S, input_2_PF, input_3_P, input_3_Q, input_3_S, input_3_PF])
        return Power


def read_acurev4100_power_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, User_Channel):
    """
    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param User_Channel:
    :return: [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [Input1_P, scale_Input1_P], [Input1_Q, scale_Input1_Q],
         [Input1_S, scale_Input1_S], [Input1_PF, scale_Input1_PF], [Input2_P, scale_Input2_P],
         [Input2_Q, scale_Input2_Q], [Input2_S, scale_Input2_S], [Input2_PF, scale_Input2_PF],
         [Input3_P, scale_Input3_P], [Input3_Q, scale_Input3_Q],
         [Input3_S, scale_Input3_S], [Input3_PF, scale_Input3_PF]~~~]
    """
    power_list = []
    Power = power_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang,
                                               User_Channel)
    Read_Phase_Power = read_phase_sys_power(Power, times=40)
    Read_Input_Power = []
    if User_Channel == "3E3p4w":
        Read_Input_Power = read_phase_input_power(Power[:12], times=40)
    if User_Channel == "1E1p2w":
        Read_Input_Power = read_phase_input_power(Power[16:], times=40)
    Phase_A_P = round(Read_Phase_Power[0], 6)
    Phase_A_Q = round(Read_Phase_Power[1], 6)
    Phase_A_S = round(Read_Phase_Power[2], 6)
    Phase_A_PF = round(Read_Phase_Power[3], 6)

    if Power[0] != 0:
        scale_Phase_A_P = abs((Phase_A_P - Power[0]) / Power[0])
    else:
        if Power[0] == Phase_A_P == 0 or Phase_A_P < 0.001:
            scale_Phase_A_P = 0
        else:
            scale_Phase_A_P = 0

    if Power[1] != 0:
        scale_Phase_A_Q = abs((Phase_A_Q - Power[1]) / Power[1])
    else:
        if Power[1] == Phase_A_Q == 0 or Phase_A_Q < 0.001:
            scale_Phase_A_Q = 0
        else:
            scale_Phase_A_Q = 0

    if Power[2] != 0:
        scale_Phase_A_S = abs((Phase_A_S - Power[2]) / Power[2])
    else:
        if Power[2] == Phase_A_S == 0 or Phase_A_S < 0.001:
            scale_Phase_A_S = 0
        else:
            scale_Phase_A_S = 0

    if Power[3] != 0:
        scale_Phase_A_PF = abs((Phase_A_PF - Power[3]) / Power[3])
    else:
        if Power[3] == Phase_A_PF == 0 or Phase_A_PF < 0.001:
            scale_Phase_A_PF = 0
        else:
            scale_Phase_A_PF = 0

    Phase_B_P = round(Read_Phase_Power[4], 6)
    Phase_B_Q = round(Read_Phase_Power[5], 6)
    Phase_B_S = round(Read_Phase_Power[6], 6)
    Phase_B_PF = round(Read_Phase_Power[7], 6)

    if Power[4] != 0:
        scale_Phase_B_P = abs((Phase_B_P - Power[4]) / Power[4])
    else:
        if Power[4] == Phase_B_P == 0 or Phase_B_P < 0.001:
            scale_Phase_B_P = 0
        else:
            scale_Phase_B_P = 0

    if Power[5] != 0:
        scale_Phase_B_Q = abs((Phase_B_Q - Power[5]) / Power[5])
    else:
        if Power[5] == Phase_B_Q == 0 or Phase_B_Q < 0.001:
            scale_Phase_B_Q = 0
        else:
            scale_Phase_B_Q = 0

    if Power[6] != 0:
        scale_Phase_B_S = abs((Phase_B_S - Power[6]) / Power[6])
    else:
        if Power[6] == Phase_B_S == 0 or Phase_B_S < 0.001:
            scale_Phase_B_S = 0
        else:
            scale_Phase_B_S = 0

    if Power[7] != 0:
        scale_Phase_B_PF = abs((Phase_B_PF - Power[7]) / Power[7])
    else:
        if Power[7] == Phase_B_PF == 0 or Phase_B_PF < 0.001:
            scale_Phase_B_PF = 0
        else:
            scale_Phase_B_PF = 0

    Phase_C_P = round(Read_Phase_Power[8], 6)
    Phase_C_Q = round(Read_Phase_Power[9], 6)
    Phase_C_S = round(Read_Phase_Power[10], 6)
    Phase_C_PF = round(Read_Phase_Power[11], 6)

    if Power[8] != 0:
        scale_Phase_C_P = abs((Phase_C_P - Power[8]) / Power[8])
    else:
        if Power[8] == Phase_C_P == 0 or Phase_C_P < 0.001:
            scale_Phase_C_P = 0
        else:
            scale_Phase_C_P = 0

    if Power[9] != 0:
        scale_Phase_C_Q = abs((Phase_C_Q - Power[9]) / Power[9])
    else:
        if Power[9] == Phase_C_Q == 0 or Phase_C_Q < 0.001:
            scale_Phase_C_Q = 0
        else:
            scale_Phase_C_Q = 0

    if Power[10] != 0:
        scale_Phase_C_S = abs((Phase_C_S - Power[10]) / Power[10])
    else:
        if Power[10] == Phase_C_S == 0 or Phase_C_S < 0.001:
            scale_Phase_C_S = 0
        else:
            scale_Phase_C_S = 0

    if Power[11] != 0:
        scale_Phase_C_PF = abs((Phase_C_PF - Power[11]) / Power[11])
    else:
        if Power[11] == Phase_C_PF == 0 or Phase_C_PF < 0.001:
            scale_Phase_C_PF = 0
        else:
            scale_Phase_C_PF = 0

    Sys_P = round(Read_Phase_Power[12], 6)
    Sys_Q = round(Read_Phase_Power[13], 6)
    Sys_S = round(Read_Phase_Power[14], 6)
    Sys_PF = round(Read_Phase_Power[15], 6)

    if Power[12] != 0:
        scale_Sys_P = abs((Sys_P - Power[12]) / Power[12])
    else:
        if Power[12] == Sys_P == 0 or Sys_P < 0.001:
            scale_Sys_P = 0
        else:
            scale_Sys_P = 0

    if Power[13] != 0:
        scale_Sys_Q = abs((Sys_Q - Power[13]) / Power[13])
    else:
        if Power[13] == Sys_Q == 0 or Sys_Q < 0.001:
            scale_Sys_Q = 0
        else:
            scale_Sys_Q = 0

    if Power[14] != 0:
        scale_Sys_S = abs((Sys_S - Power[14]) / Power[14])
    else:
        if Power[14] == Sys_S == 0 or Sys_S < 0.001:
            scale_Sys_S = 0
        else:
            scale_Sys_S = 0

    if Power[15] != 0:
        scale_Sys_PF = abs((Sys_PF - Power[15]) / Power[15])
    else:
        if Power[15] == Sys_PF == 0 or Sys_PF < 0.001:
            scale_Sys_PF = 0
        else:
            scale_Sys_PF = 0
    if User_Channel == '3E3p4w':
        Input1_P = round(Read_Input_Power[0], 6)
        Input1_Q = round(Read_Input_Power[1], 6)
        Input1_S = round(Read_Input_Power[2], 6)
        Input1_PF = round(Read_Input_Power[3], 6)

        if Power[0] != 0:
            scale_Input1_P = abs((Input1_P - Power[0]) / Power[0])
        else:
            if Power[0] == Input1_P == 0 or Input1_P < 0.001:
                scale_Input1_P = 0
            else:
                scale_Input1_P = 0

        if Power[1] != 0:
            scale_Input1_Q = abs((Input1_Q - Power[1]) / Power[1])
        else:
            if Power[1] == Input1_Q == 0 or Input1_Q < 0.001:
                scale_Input1_Q = 0
            else:
                scale_Input1_Q = 0

        if Power[2] != 0:
            scale_Input1_S = abs((Input1_S - Power[2]) / Power[2])
        else:
            if Power[2] == Input1_S == 0 or Input1_S < 0.001:
                scale_Input1_S = 0
            else:
                scale_Input1_S = 0
        if Power[3] != 0:
            scale_Input1_PF = abs((Input1_PF - Power[3]) / Power[3])
        else:
            if Power[3] == Input1_PF == 0 or Input1_PF < 0.001:
                scale_Input1_PF = 0
            else:
                scale_Input1_PF = 0

        Input2_P = round(Read_Input_Power[4], 6)
        Input2_Q = round(Read_Input_Power[5], 6)
        Input2_S = round(Read_Input_Power[6], 6)
        Input2_PF = round(Read_Input_Power[7], 6)

        if Power[4] != 0:
            scale_Input2_P = abs((Input2_P - Power[4]) / Power[4])
        else:
            if Power[4] == Input2_P == 0 or Input2_P < 0.001:
                scale_Input2_P = 0
            else:
                scale_Input2_P = 0

        if Power[5] != 0:
            scale_Input2_Q = abs((Input2_Q - Power[5]) / Power[5])
        else:
            if Power[5] == Input2_Q == 0 or Input2_Q < 0.001:
                scale_Input2_Q = 0
            else:
                scale_Input2_Q = 0

        if Power[6] != 0:
            scale_Input2_S = abs((Input2_S - Power[6]) / Power[6])
        else:
            if Power[6] == Input2_S == 0 or Input2_S < 0.001:
                scale_Input2_S = 0
            else:
                scale_Input2_S = 0

        if Power[7] != 0:
            scale_Input2_PF = abs((Input2_PF - Power[7]) / Power[7])
        else:
            if Power[7] == Input2_PF == 0 or Input2_PF < 0.001:
                scale_Input2_PF = 0
            else:
                scale_Input2_PF = 0

        Input3_P = round(Read_Input_Power[8], 6)
        Input3_Q = round(Read_Input_Power[9], 6)
        Input3_S = round(Read_Input_Power[10], 6)
        Input3_PF = round(Read_Input_Power[11], 6)

        if Power[8] != 0:
            scale_Input3_P = abs((Input3_P - Power[8]) / Power[8])
        else:
            if Power[8] == Input3_P == 0 or Input3_P < 0.001:
                scale_Input3_P = 0
            else:
                scale_Input3_P = 0

        if Power[9] != 0:
            scale_Input3_Q = abs((Input3_Q - Power[9]) / Power[9])
        else:
            if Power[9] == Input3_Q == 0 or Input3_Q < 0.001:
                scale_Input3_Q = 0
            else:
                scale_Input3_Q = 0

        if Power[10] != 0:
            scale_Input3_S = abs((Input3_S - Power[10]) / Power[10])
        else:
            if Power[10] == Input3_S == 0 or Input3_S < 0.001:
                scale_Input3_S = 0
            else:
                scale_Input3_S = 0

        if Power[11] != 0:
            scale_Input3_PF = abs((Input3_PF - Power[11]) / Power[11])
        else:
            if Power[11] == Input3_PF == 0 or Input3_PF < 0.001:
                scale_Input3_PF = 0
            else:
                scale_Input3_PF = 0
    if User_Channel == '1E1p2w':
        Input1_P = round(Read_Input_Power[0], 6)
        Input1_Q = round(Read_Input_Power[1], 6)
        Input1_S = round(Read_Input_Power[2], 6)
        Input1_PF = round(Read_Input_Power[3], 6)

        if Power[16] != 0:
            scale_Input1_P = abs((Input1_P - Power[16]) / Power[16])
        else:
            if Power[16] == Input1_P == 0 or Input1_P < 0.001:
                scale_Input1_P = 0
            else:
                scale_Input1_P = 0

        if Power[17] != 0:
            scale_Input1_Q = abs((Input1_Q - Power[17]) / Power[17])
        else:
            if Power[17] == Input1_Q == 0 or Input1_Q < 0.001:
                scale_Input1_Q = 0
            else:
                scale_Input1_Q = 0

        if Power[18] != 0:
            scale_Input1_S = abs((Input1_S - Power[18]) / Power[18])
        else:
            if Power[18] == Input1_S == 0 or Input1_S < 0.001:
                scale_Input1_S = 0
            else:
                scale_Input1_S = 0

        if Power[19] != 0:
            scale_Input1_PF = abs((Input1_PF - Power[19]) / Power[19])
        else:
            if Power[19] == Input1_PF == 0 or Input1_PF < 0.001:
                scale_Input1_PF = 0
            else:
                scale_Input1_PF = 0

        Input2_P = round(Read_Input_Power[4], 6)
        Input2_Q = round(Read_Input_Power[5], 6)
        Input2_S = round(Read_Input_Power[6], 6)
        Input2_PF = round(Read_Input_Power[7], 6)

        if Power[20] != 0:
            scale_Input2_P = abs((Input2_P - Power[20]) / Power[20])
        else:
            if Power[20] == Input2_P == 0 or Input2_P < 0.001:
                scale_Input2_P = 0
            else:
                scale_Input2_P = 0

        if Power[21] != 0:
            scale_Input2_Q = abs((Input2_Q - Power[21]) / Power[21])
        else:
            if Power[21] == Input2_Q == 0 or Input2_Q < 0.001:
                scale_Input2_Q = 0
            else:
                scale_Input2_Q = 0

        if Power[22] != 0:
            scale_Input2_S = abs((Input2_S - Power[22]) / Power[22])
        else:
            if Power[22] == Input2_S == 0 or Input2_S < 0.001:
                scale_Input2_S = 0
            else:
                scale_Input2_S = 0

        if Power[23] != 0:
            scale_Input2_PF = abs((Input2_PF - Power[23]) / Power[23])
        else:
            if Power[23] == Input2_PF == 0 or Input2_PF < 0.001:
                scale_Input2_PF = 0
            else:
                scale_Input2_PF = 0

        Input3_P = round(Read_Input_Power[8], 6)
        Input3_Q = round(Read_Input_Power[9], 6)
        Input3_S = round(Read_Input_Power[10], 6)
        Input3_PF = round(Read_Input_Power[11], 6)

        if Power[24] != 0:
            scale_Input3_P = abs((Input3_P - Power[24]) / Power[24])
        else:
            if Power[24] == Input3_P == 0 or Input3_P < 0.001:
                scale_Input3_P = 0
            else:
                scale_Input3_P = 0

        if Power[25] != 0:
            scale_Input3_Q = abs((Input3_Q - Power[25]) / Power[25])
        else:
            if Power[25] == Input3_Q == 0 or Input3_Q < 0.001:
                scale_Input3_Q = 0
            else:
                scale_Input3_Q = 0

        if Power[26] != 0:
            scale_Input3_S = abs((Input3_S - Power[26]) / Power[26])
        else:
            if Power[26] == Input3_S == 0 or Input3_S < 0.001:
                scale_Input3_S = 0
            else:
                scale_Input3_S = 0

        if Power[27] != 0:
            scale_Input3_PF = abs((Input3_PF - Power[27]) / Power[27])
        else:
            if Power[27] == Input3_PF == 0 or Input3_PF < 0.001:
                scale_Input3_PF = 0
            else:
                scale_Input3_PF = 0

    power_list.extend(
        [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [Input1_P, scale_Input1_P], [Input1_Q, scale_Input1_Q],
         [Input1_S, scale_Input1_S], [Input1_PF, scale_Input1_PF], [Input2_P, scale_Input2_P],
         [Input2_Q, scale_Input2_Q], [Input2_S, scale_Input2_S], [Input2_PF, scale_Input2_PF],
         [Input3_P, scale_Input3_P], [Input3_Q, scale_Input3_Q],
         [Input3_S, scale_Input3_S], [Input3_PF, scale_Input3_PF]])

    if User_Channel == '3E3p4w':
        User1_P = read_user_channel_1_active_power(Power[12], times=40)
        User1_Q = read_user_channel_1_reactive_power(Power[13], times=40)
        User1_S = read_user_channel_1_apparent_power(Power[14], times=40)
        User1_PF = read_user_channel_1_power_factor(Power[15], times=40)

        if Power[12] != 0:
            scale_User1_P = abs((User1_P - Power[12]) / Power[12])
        else:
            if Power[12] == User1_P == 0 or User1_P < 0.001:
                scale_User1_P = 0
            else:
                scale_User1_P = 0

        if Power[13] != 0:
            scale_User1_Q = abs((User1_Q - Power[13]) / Power[13])
        else:
            if Power[13] == User1_Q == 0 or User1_Q < 0.001:
                scale_User1_Q = 0
            else:
                scale_User1_Q = 0

        if Power[14] != 0:
            scale_User1_S = abs((User1_S - Power[14]) / Power[14])
        else:
            if Power[14] == User1_S == 0 or User1_S < 0.001:
                scale_User1_S = 0
            else:
                scale_User1_S = 0

        if Power[15] != 0:
            scale_User1_PF = abs((User1_PF - Power[15]) / Power[15])
        else:
            if Power[15] == User1_PF == 0 or User1_PF < 0.001:
                scale_User1_PF = 0
            else:
                scale_User1_PF = 0

        power_list.extend(
            [[User1_P, scale_User1_P], [User1_Q, scale_User1_Q], [User1_S, scale_User1_S], [User1_PF, scale_User1_PF]])
    if User_Channel == '1E1p2w':
        Read_User_Power_list = read_user_power([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], times=40)
        User1_P = Read_User_Power_list[0]
        User1_Q = Read_User_Power_list[1]
        User1_S = Read_User_Power_list[2]
        User1_PF = Read_User_Power_list[3]

        if User1_P == 0 or User1_P < 0.001:
            scale_User1_P = 0
        else:
            scale_User1_P = 0

        if User1_Q == 0 or User1_Q < 0.001:
            scale_User1_Q = 0
        else:
            scale_User1_Q = 0

        if User1_S == 0 or User1_S < 0.001:
            scale_User1_S = 0
        else:
            scale_User1_S = 0

        if User1_PF == 0 or User1_PF < 0.001:
            scale_User1_PF = 0
        else:
            scale_User1_PF = 0

        User2_P = Read_User_Power_list[4]
        User2_Q = Read_User_Power_list[5]
        User2_S = Read_User_Power_list[6]
        User2_PF = Read_User_Power_list[7]

        if User2_P == 0 or User2_P < 0.001:
            scale_User2_P = 0
        else:
            scale_User2_P = 0

        if User2_Q == 0 or User2_Q < 0.001:
            scale_User2_Q = 0
        else:
            scale_User2_Q = 0

        if User2_S == 0 or User2_S < 0.001:
            scale_User2_S = 0
        else:
            scale_User2_S = 0

        if User2_PF == 0 or User2_PF < 0.001:
            scale_User2_PF = 0
        else:
            scale_User2_PF = 0

        User3_P = Read_User_Power_list[8]
        User3_Q = Read_User_Power_list[9]
        User3_S = Read_User_Power_list[10]
        User3_PF = Read_User_Power_list[11]

        if User3_P == 0 or User3_P < 0.001:
            scale_User3_P = 0
        else:
            scale_User3_P = 0

        if User3_Q == 0 or User3_Q < 0.001:
            scale_User3_Q = 0
        else:
            scale_User3_Q = 0

        if User3_S == 0 or User3_S < 0.001:
            scale_User3_S = 0
        else:
            scale_User3_S = 0

        if User3_PF == 0 or User3_PF < 0.001:
            scale_User3_PF = 0
        else:
            scale_User3_PF = 0
        power_list.extend(
            [[User1_P, scale_User1_P], [User1_Q, scale_User1_Q], [User1_S, scale_User1_S], [User1_PF, scale_User1_PF],
             [User2_P, scale_User2_P], [User2_Q, scale_User2_Q], [User2_S, scale_User2_S], [User2_PF, scale_User2_PF],
             [User3_P, scale_User3_P], [User3_Q, scale_User3_Q], [User3_S, scale_User3_S], [User3_PF, scale_User3_PF]])
    return power_list


def energy_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time,
                                        Service):
    '''

    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param Time:
    :return: Phase_A_P_E_import, Phase_A_P_E_export, Phase_A_P_E_net, Phase_A_P_E_total, Phase_A_Q_E_import,
         Phase_A_Q_E_export, Phase_A_Q_E_net, Phase_A_Q_E_total, Phase_A_S_E, Phase_B_P_E_import, Phase_B_P_E_export,
         Phase_B_P_E_net, Phase_B_P_E_total, Phase_B_Q_E_import, Phase_B_Q_E_export, Phase_B_Q_E_net, Phase_B_Q_E_total,
         Phase_B_S_E, Phase_C_P_E_import, Phase_C_P_E_export, Phase_C_P_E_net, Phase_C_P_E_total, Phase_C_Q_E_import,
         Phase_C_Q_E_export, Phase_C_Q_E_net, Phase_C_Q_E_total, Phase_C_S_E, System_P_E_import, System_P_E_export,
         System_P_E_net, System_P_E_total, System_Q_E_import, System_Q_E_export, System_Q_E_net, System_Q_E_total,
         System_S_E
    '''
    if Service == "3E3p4w":
        Energy_Standard_Value_list = []
        Power_standard_value_list = power_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
                                                                       Ia_ang,
                                                                       Ib_ang, Ic_ang, Service)
        Phase_A_P = Power_standard_value_list[0]
        Phase_A_Q = Power_standard_value_list[1]
        Phase_A_S = Power_standard_value_list[2]
        if Phase_A_P > 0:
            Phase_A_P_E_import = abs(Phase_A_P * (Time / 60))
            Phase_A_P_E_export = 0
        else:
            Phase_A_P_E_import = 0
            Phase_A_P_E_export = abs(Phase_A_P * (Time / 60))
        Phase_A_P_E_net = Phase_A_P_E_import - Phase_A_P_E_export
        Phase_A_P_E_total = Phase_A_P_E_import + Phase_A_P_E_export

        if Phase_A_Q > 0:
            Phase_A_Q_E_import = abs(Phase_A_Q * (Time / 60))
            Phase_A_Q_E_export = 0
        else:
            Phase_A_Q_E_import = 0
            Phase_A_Q_E_export = abs(Phase_A_Q * (Time / 60))
        Phase_A_Q_E_net = Phase_A_Q_E_import - Phase_A_Q_E_export
        Phase_A_Q_E_total = Phase_A_Q_E_import + Phase_A_Q_E_export

        Phase_A_S_E = Phase_A_S * (Time / 60)

        Phase_B_P = Power_standard_value_list[4]
        Phase_B_Q = Power_standard_value_list[5]
        Phase_B_S = Power_standard_value_list[6]
        if Phase_B_P > 0:
            Phase_B_P_E_import = abs(Phase_B_P * (Time / 60))
            Phase_B_P_E_export = 0
        else:
            Phase_B_P_E_import = 0
            Phase_B_P_E_export = abs(Phase_B_P * (Time / 60))
        Phase_B_P_E_net = Phase_B_P_E_import - Phase_B_P_E_export
        Phase_B_P_E_total = Phase_B_P_E_import + Phase_B_P_E_export

        if Phase_B_Q > 0:
            Phase_B_Q_E_import = abs(Phase_B_Q * (Time / 60))
            Phase_B_Q_E_export = 0
        else:
            Phase_B_Q_E_import = 0
            Phase_B_Q_E_export = abs(Phase_B_Q * (Time / 60))
        Phase_B_Q_E_net = Phase_B_Q_E_import - Phase_B_Q_E_export
        Phase_B_Q_E_total = Phase_B_Q_E_import + Phase_B_Q_E_export

        Phase_B_S_E = Phase_B_S * (Time / 60)

        Phase_C_P = Power_standard_value_list[8]
        Phase_C_Q = Power_standard_value_list[9]
        Phase_C_S = Power_standard_value_list[10]
        if Phase_C_P > 0:
            Phase_C_P_E_import = abs(Phase_C_P * (Time / 60))
            Phase_C_P_E_export = 0
        else:
            Phase_C_P_E_import = 0
            Phase_C_P_E_export = abs(Phase_C_P * (Time / 60))
        Phase_C_P_E_net = Phase_C_P_E_import - Phase_C_P_E_export
        Phase_C_P_E_total = Phase_C_P_E_import + Phase_C_P_E_export

        if Phase_C_Q > 0:
            Phase_C_Q_E_import = abs(Phase_C_Q * (Time / 60))
            Phase_C_Q_E_export = 0
        else:
            Phase_C_Q_E_import = 0
            Phase_C_Q_E_export = abs(Phase_C_Q * (Time / 60))
        Phase_C_Q_E_net = Phase_C_Q_E_import - Phase_C_Q_E_export
        Phase_C_Q_E_total = Phase_C_Q_E_import + Phase_C_Q_E_export

        Phase_C_S_E = Phase_C_S * (Time / 60)

        System_P_E_import = Phase_A_P_E_import + Phase_B_P_E_import + Phase_C_P_E_import
        System_P_E_export = Phase_A_P_E_export + Phase_B_P_E_export + Phase_C_P_E_export
        System_P_E_net = Phase_A_P_E_net + Phase_B_P_E_net + Phase_C_P_E_net
        System_P_E_total = Phase_A_P_E_total + Phase_B_P_E_total + Phase_C_P_E_total

        System_Q_E_import = Phase_A_Q_E_import + Phase_B_Q_E_import + Phase_C_Q_E_import
        System_Q_E_export = Phase_A_Q_E_export + Phase_B_Q_E_export + Phase_C_Q_E_export
        System_Q_E_net = Phase_A_Q_E_net + Phase_B_Q_E_net + Phase_C_Q_E_net
        System_Q_E_total = Phase_A_Q_E_total + Phase_B_Q_E_total + Phase_C_Q_E_total

        System_S_E = Phase_A_S_E + Phase_B_S_E + Phase_C_S_E

        Energy_Standard_Value_list.extend(
            [Phase_A_P_E_import, Phase_A_P_E_export, Phase_A_P_E_net, Phase_A_P_E_total, Phase_A_Q_E_import,
             Phase_A_Q_E_export, Phase_A_Q_E_net, Phase_A_Q_E_total, Phase_A_S_E, Phase_B_P_E_import,
             Phase_B_P_E_export,
             Phase_B_P_E_net, Phase_B_P_E_total, Phase_B_Q_E_import, Phase_B_Q_E_export, Phase_B_Q_E_net,
             Phase_B_Q_E_total,
             Phase_B_S_E, Phase_C_P_E_import, Phase_C_P_E_export, Phase_C_P_E_net, Phase_C_P_E_total,
             Phase_C_Q_E_import,
             Phase_C_Q_E_export, Phase_C_Q_E_net, Phase_C_Q_E_total, Phase_C_S_E, System_P_E_import, System_P_E_export,
             System_P_E_net, System_P_E_total, System_Q_E_import, System_Q_E_export, System_Q_E_net, System_Q_E_total,
             System_S_E])
        return Energy_Standard_Value_list
    if Service == "1E1p2w":
        Energy_Standard_Value_list = []
        Power_standard_value_list = power_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
                                                                       Ia_ang,
                                                                       Ib_ang, Ic_ang, Service)
        Phase_A_P = Power_standard_value_list[0]
        Phase_A_Q = Power_standard_value_list[1]
        Phase_A_S = Power_standard_value_list[2]
        if Phase_A_P > 0:
            Phase_A_P_E_import = abs(Phase_A_P * (Time / 60))
            Phase_A_P_E_export = 0
        else:
            Phase_A_P_E_import = 0
            Phase_A_P_E_export = abs(Phase_A_P * (Time / 60))
        Phase_A_P_E_net = Phase_A_P_E_import - Phase_A_P_E_export
        Phase_A_P_E_total = Phase_A_P_E_import + Phase_A_P_E_export

        if Phase_A_Q > 0:
            Phase_A_Q_E_import = abs(Phase_A_Q * (Time / 60))
            Phase_A_Q_E_export = 0
        else:
            Phase_A_Q_E_import = 0
            Phase_A_Q_E_export = abs(Phase_A_Q * (Time / 60))
        Phase_A_Q_E_net = Phase_A_Q_E_import - Phase_A_Q_E_export
        Phase_A_Q_E_total = Phase_A_Q_E_import + Phase_A_Q_E_export

        Phase_A_S_E = Phase_A_S * (Time / 60)

        Phase_B_P = Power_standard_value_list[4]
        Phase_B_Q = Power_standard_value_list[5]
        Phase_B_S = Power_standard_value_list[6]
        if Phase_B_P > 0:
            Phase_B_P_E_import = abs(Phase_B_P * (Time / 60))
            Phase_B_P_E_export = 0
        else:
            Phase_B_P_E_import = 0
            Phase_B_P_E_export = abs(Phase_B_P * (Time / 60))
        Phase_B_P_E_net = Phase_B_P_E_import - Phase_B_P_E_export
        Phase_B_P_E_total = Phase_B_P_E_import + Phase_B_P_E_export

        if Phase_B_Q > 0:
            Phase_B_Q_E_import = abs(Phase_B_Q * (Time / 60))
            Phase_B_Q_E_export = 0
        else:
            Phase_B_Q_E_import = 0
            Phase_B_Q_E_export = abs(Phase_B_Q * (Time / 60))
        Phase_B_Q_E_net = Phase_B_Q_E_import - Phase_B_Q_E_export
        Phase_B_Q_E_total = Phase_B_Q_E_import + Phase_B_Q_E_export

        Phase_B_S_E = Phase_B_S * (Time / 60)

        Phase_C_P = Power_standard_value_list[8]
        Phase_C_Q = Power_standard_value_list[9]
        Phase_C_S = Power_standard_value_list[10]
        if Phase_C_P > 0:
            Phase_C_P_E_import = abs(Phase_C_P * (Time / 60))
            Phase_C_P_E_export = 0
        else:
            Phase_C_P_E_import = 0
            Phase_C_P_E_export = abs(Phase_C_P * (Time / 60))
        Phase_C_P_E_net = Phase_C_P_E_import - Phase_C_P_E_export
        Phase_C_P_E_total = Phase_C_P_E_import + Phase_C_P_E_export

        if Phase_C_Q > 0:
            Phase_C_Q_E_import = abs(Phase_C_Q * (Time / 60))
            Phase_C_Q_E_export = 0
        else:
            Phase_C_Q_E_import = 0
            Phase_C_Q_E_export = abs(Phase_C_Q * (Time / 60))
        Phase_C_Q_E_net = Phase_C_Q_E_import - Phase_C_Q_E_export
        Phase_C_Q_E_total = Phase_C_Q_E_import + Phase_C_Q_E_export

        Phase_C_S_E = Phase_C_S * (Time / 60)

        input_1_P = Power_standard_value_list[16]
        input_1_Q = Power_standard_value_list[17]
        input_1_S = Power_standard_value_list[18]
        if input_1_P > 0:
            input_1_P_E_import = abs(input_1_P * (Time / 60))
            input_1_P_E_export = 0
        else:
            input_1_P_E_import = 0
            input_1_P_E_export = abs(input_1_P * (Time / 60))
        input_1_P_E_net = input_1_P_E_import - input_1_P_E_export
        input_1_P_E_total = input_1_P_E_import + input_1_P_E_export

        if input_1_Q > 0:
            input_1_Q_E_import = abs(input_1_Q * (Time / 60))
            input_1_Q_E_export = 0
        else:
            input_1_Q_E_import = 0
            input_1_Q_E_export = abs(input_1_Q * (Time / 60))
        input_1_Q_E_net = input_1_Q_E_import - input_1_Q_E_export
        input_1_Q_E_total = input_1_Q_E_import + input_1_Q_E_export

        input_1_S_E = input_1_S * (Time / 60)

        input_2_P = Power_standard_value_list[20]
        input_2_Q = Power_standard_value_list[21]
        input_2_S = Power_standard_value_list[22]
        if input_2_P > 0:
            input_2_P_E_import = abs(input_2_P * (Time / 60))
            input_2_P_E_export = 0
        else:
            input_2_P_E_import = 0
            input_2_P_E_export = abs(input_2_P * (Time / 60))
        input_2_P_E_net = input_2_P_E_import - input_2_P_E_export
        input_2_P_E_total = input_2_P_E_import + input_2_P_E_export

        if input_2_Q > 0:
            input_2_Q_E_import = abs(input_2_Q * (Time / 60))
            input_2_Q_E_export = 0
        else:
            input_2_Q_E_import = 0
            input_2_Q_E_export = abs(input_2_Q * (Time / 60))
        input_2_Q_E_net = input_2_Q_E_import - input_2_Q_E_export
        input_2_Q_E_total = input_2_Q_E_import + input_2_Q_E_export

        input_2_S_E = input_2_S * (Time / 60)

        input_3_P = Power_standard_value_list[24]
        input_3_Q = Power_standard_value_list[25]
        input_3_S = Power_standard_value_list[26]
        if input_3_P > 0:
            input_3_P_E_import = abs(input_3_P * (Time / 60))
            input_3_P_E_export = 0
        else:
            input_3_P_E_import = 0
            input_3_P_E_export = abs(input_3_P * (Time / 60))
        input_3_P_E_net = input_3_P_E_import - input_3_P_E_export
        input_3_P_E_total = input_3_P_E_import + input_3_P_E_export

        if input_3_Q > 0:
            input_3_Q_E_import = abs(input_3_Q * (Time / 60))
            input_3_Q_E_export = 0
        else:
            input_3_Q_E_import = 0
            input_3_Q_E_export = abs(input_3_Q * (Time / 60))
        input_3_Q_E_net = input_3_Q_E_import - input_3_Q_E_export
        input_3_Q_E_total = input_3_Q_E_import + input_3_Q_E_export

        input_3_S_E = input_3_S * (Time / 60)

        System_P_E_import = Phase_A_P_E_import
        System_P_E_export = Phase_A_P_E_export
        System_P_E_net = Phase_A_P_E_net
        System_P_E_total = Phase_A_P_E_total

        System_Q_E_import = Phase_A_Q_E_import
        System_Q_E_export = Phase_A_Q_E_export
        System_Q_E_net = Phase_A_Q_E_net
        System_Q_E_total = Phase_A_Q_E_total

        System_S_E = Phase_A_S_E

        Energy_Standard_Value_list.extend(
            [Phase_A_P_E_import, Phase_A_P_E_export, Phase_A_P_E_net, Phase_A_P_E_total, Phase_A_Q_E_import,
             Phase_A_Q_E_export, Phase_A_Q_E_net, Phase_A_Q_E_total, Phase_A_S_E, Phase_B_P_E_import,
             Phase_B_P_E_export,
             Phase_B_P_E_net, Phase_B_P_E_total, Phase_B_Q_E_import, Phase_B_Q_E_export, Phase_B_Q_E_net,
             Phase_B_Q_E_total,
             Phase_B_S_E, Phase_C_P_E_import, Phase_C_P_E_export, Phase_C_P_E_net, Phase_C_P_E_total,
             Phase_C_Q_E_import,
             Phase_C_Q_E_export, Phase_C_Q_E_net, Phase_C_Q_E_total, Phase_C_S_E, System_P_E_import, System_P_E_export,
             System_P_E_net, System_P_E_total, System_Q_E_import, System_Q_E_export, System_Q_E_net, System_Q_E_total,
             System_S_E, input_1_P_E_import, input_1_P_E_export, input_1_P_E_net, input_1_P_E_total, input_1_Q_E_import,
             input_1_Q_E_export, input_1_Q_E_net, input_1_Q_E_total, input_1_S_E, input_2_P_E_import,
             input_2_P_E_export, input_2_P_E_net, input_2_P_E_total, input_2_Q_E_import, input_2_Q_E_export,
             input_2_Q_E_net, input_2_Q_E_total, input_2_S_E, input_3_P_E_import, input_3_P_E_export, input_3_P_E_net,
             input_3_P_E_total, input_3_Q_E_import, input_3_Q_E_export, input_3_Q_E_net, input_3_Q_E_total,
             input_3_S_E])
        return Energy_Standard_Value_list


def read_energy_scale_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time, Service):
    '''

    :param Va:
    :param Vb:
    :param Vc:
    :param Ia:
    :param Ib:
    :param Ic:
    :param Va_ang:
    :param Vb_ang:
    :param Vc_ang:
    :param Ia_ang:
    :param Ib_ang:
    :param Ic_ang:
    :param Time:
    :param Service:
    :return:
    :Read_Energy_list:4100_Energy_list
    :Energy_scale_list
    '''
    Energy_Standard_Value_list = energy_standard_value_calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
                                                                     Ia_ang,
                                                                     Ib_ang, Ic_ang, Time, Service)
    Phase_A_Energy_list = read_phase_a_energy()
    Phase_B_Energy_list = read_phase_b_energy()
    Phase_C_Energy_list = read_phase_c_energy()
    System_Energy_list = read_system_energy()
    Input_Channel_1_Energy = read_input_channel_1_energy()
    Input_Channel_2_Energy = read_input_channel_2_energy()
    Input_Channel_3_Energy = read_input_channel_3_energy()
    User_Channel_1_Energy = read_user_channel_1_energy()
    User_Channel_2_Energy = read_user_channel_2_energy()
    User_Channel_3_Energy = read_user_channel_3_energy()
    Read_Energy_list = []
    Read_Energy_list.extend(
        Phase_A_Energy_list + Phase_B_Energy_list + Phase_C_Energy_list + System_Energy_list + Input_Channel_1_Energy +
        Input_Channel_2_Energy + Input_Channel_3_Energy + User_Channel_1_Energy + User_Channel_2_Energy + User_Channel_3_Energy)
    Energy_scale_list = []
    for i in range(36):
        if Energy_Standard_Value_list[i] != 0:
            Energy_scale = abs((Read_Energy_list[i] - Energy_Standard_Value_list[i]) / Energy_Standard_Value_list[i])
            Energy_scale_list.append(Energy_scale)
        elif Energy_Standard_Value_list[i] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
            Energy_scale = 0
            Energy_scale_list.append(Energy_scale)
        else:
            Energy_scale = 0
            Energy_scale_list.append(Energy_scale)

    if Service == '3E3p4w':
        for i in range(36, 63):
            if Energy_Standard_Value_list[i - 36] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i - 36]) / Energy_Standard_Value_list[i - 36])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i - 36] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)

        for i in range(63, 72):
            if Energy_Standard_Value_list[i - 36] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i - 36]) / Energy_Standard_Value_list[i - 36])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i - 36] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
        for i in range(72, 90):
            if Read_Energy_list[i] == 0:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
    if Service == '1E1p2w':
        for i in range(36, 63):
            if Energy_Standard_Value_list[i] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i]) / Energy_Standard_Value_list[i])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
        # for i in range(63, 90):
        #     if Energy_Standard_Value_list[i - 27] != 0:
        #         Energy_scale = abs(
        #             (Read_Energy_list[i] - Energy_Standard_Value_list[i - 27]) / Energy_Standard_Value_list[i - 27])
        #         Energy_scale_list.append(Energy_scale)
        #     elif Energy_Standard_Value_list[i - 27] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
        #         Energy_scale = 0
        #         Energy_scale_list.append(Energy_scale)
        #     else:
        #         Energy_scale = 'null'
        #         Energy_scale_list.append(Energy_scale)
        for i in range(63, 90):
            if Read_Energy_list[i] == 0:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
    return Read_Energy_list, Energy_scale_list


def wait_minutes(wait_times):
    """等待指定分钟后输出"""
    time.sleep(wait_times * 60)  # 转换成秒
    logging.info(f"等待 {wait_times} 分钟结束！")


def hold_rs485_connect(hold_time):
    """每 5 分钟读取一次 Modbus 数据"""
    t = int((hold_time / 3) - 1)  # 计算读取次数
    for i in range(t):
        time.sleep(180)  # 3 分钟
        value = ModbusClient.read_measurement(address=real_time_addr['System Frequency']['Start(Dec)'],
                                              count=real_time_addr['System Frequency']['Reg'], slave=1)
        logging.info(f"第 {i + 1} 次读取数据: {value},RS485 连接正常")


def calculate_k_factor(base_current, harmonics):
    """
    计算K系数
    :param base_current: 基波电流 (I1)
    :param harmonics: 一个字典，键为谐波次数（h），值为该次谐波占基波电流的百分比（0~100）
    :return: K系数
    """
    numerator = 0
    denominator = 0

    # 添加基波
    numerator += (base_current ** 2) * (1 ** 2)
    denominator += base_current ** 2

    # 处理其他谐波
    for h, percent in harmonics.items():
        ih = (percent / 100) * base_current
        numerator += (ih ** 2) * (h ** 2)
        denominator += ih ** 2

    k_factor = numerator / denominator
    return k_factor


def e2_3w_1p_pf_calculate_angle(PF):
    if PF == 1:
        return 180, 0, 0, 0, 180, 0
    if PF == '0.5L':
        return 180, 0, 0, 0, 120, 300
    if PF == '0.8C':
        return 180, 0, 0, 0, 216.87, 36.87


def power_level(PF, In):
    if PF == 1 and In == "0.01In":
        return 0.4
    if PF == 1 and In in ("0.05In", "0.5Imax", "Imax"):
        return 0.2
    if PF in ('0.5L', '0.8C') and In == "0.01Imax":
        return 0.49
    if PF in ('0.5L', '0.8C') and In == "0.02In":
        return 0.49
    if PF in ('0.5L', '0.8C') and In in ("0.1In", "Imax"):
        return 0.21


def e2_3w_1p_power_standard_value(Va, Vc, I1, I2, Va_ang, I1_ang, Vc_ang, I2_ang):
    Power = []
    Phase_A_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_S = round(apparent_power_calculate(Va, I1), 10)
    Phase_A_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    Phase_B_P = 0
    Phase_B_Q = 0
    Phase_B_S = 0
    Phase_B_PF = 1

    Phase_C_P = round(active_power_calculate(Vc, I2, Vc_ang, I2_ang), 10)
    Phase_C_Q = round(reactive_power_calculate(Vc, I2, Vc_ang, I2_ang), 10)
    Phase_C_S = round(apparent_power_calculate(Vc, I2), 10)
    Phase_C_PF = round(power_factor_calculate(Vc_ang, I2_ang), 10)

    Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
    Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
    Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend(
        [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
         Phase_C_Q,
         Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
    return Power


def e2_3w_1p_power(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    power_list = []
    Power = e2_3w_1p_power_standard_value(Va, Vc, I1, I2, Va_ang, I1_ang, Vc_ang, I2_ang)
    Read_Phase_Power = read_phase_sys_power(Power, times=10)
    User1_P = read_user_channel_1_active_power(Power[12], times=10)
    User1_Q = read_user_channel_1_reactive_power(Power[13], times=10)
    User1_S = read_user_channel_1_apparent_power(Power[14], times=10)
    User1_PF = read_user_channel_1_power_factor(Power[15], times=10)

    Phase_A_P = round(Read_Phase_Power[0], 6)
    Phase_A_Q = round(Read_Phase_Power[1], 6)
    Phase_A_S = round(Read_Phase_Power[2], 6)
    Phase_A_PF = round(Read_Phase_Power[3], 6)

    if Power[0] != 0:
        scale_Phase_A_P = abs((Phase_A_P - Power[0]) / Power[0])
    else:
        if Power[0] == Phase_A_P == 0 or Phase_A_P < 0.001:
            scale_Phase_A_P = 0
        else:
            scale_Phase_A_P = 'null'

    if Power[1] != 0:
        scale_Phase_A_Q = abs((Phase_A_Q - Power[1]) / Power[1])
    else:
        if Power[1] == Phase_A_Q == 0 or Phase_A_Q < 0.001:
            scale_Phase_A_Q = 0
        else:
            scale_Phase_A_Q = 'null'

    if Power[2] != 0:
        scale_Phase_A_S = abs((Phase_A_S - Power[2]) / Power[2])
    else:
        if Power[2] == Phase_A_S == 0 or Phase_A_S < 0.001:
            scale_Phase_A_S = 0
        else:
            scale_Phase_A_S = 'null'

    if Power[3] != 0:
        scale_Phase_A_PF = abs((Phase_A_PF - Power[3]) / Power[3])
    else:
        if Power[3] == Phase_A_PF == 0 or Phase_A_PF < 0.001:
            scale_Phase_A_PF = 0
        else:
            scale_Phase_A_PF = 'null'

    Phase_B_P = round(Read_Phase_Power[4], 6)
    Phase_B_Q = round(Read_Phase_Power[5], 6)
    Phase_B_S = round(Read_Phase_Power[6], 6)
    Phase_B_PF = round(Read_Phase_Power[7], 6)

    if Power[4] != 0:
        scale_Phase_B_P = abs((Phase_B_P - Power[4]) / Power[4])
    else:
        if Power[4] == Phase_B_P == 0 or Phase_B_P < 0.001:
            scale_Phase_B_P = 0
        else:
            scale_Phase_B_P = 'null'

    if Power[5] != 0:
        scale_Phase_B_Q = abs((Phase_B_Q - Power[5]) / Power[5])
    else:
        if Power[5] == Phase_B_Q == 0 or Phase_B_Q < 0.001:
            scale_Phase_B_Q = 0
        else:
            scale_Phase_B_Q = 'null'

    if Power[6] != 0:
        scale_Phase_B_S = abs((Phase_B_S - Power[6]) / Power[6])
    else:
        if Power[6] == Phase_B_S == 0 or Phase_B_S < 0.001:
            scale_Phase_B_S = 0
        else:
            scale_Phase_B_S = 'null'

    if Power[7] != 0:
        scale_Phase_B_PF = abs((Phase_B_PF - Power[7]) / Power[7])
    else:
        if Power[7] == Phase_B_PF == 0 or Phase_B_PF < 0.001:
            scale_Phase_B_PF = 0
        else:
            scale_Phase_B_PF = 'null'

    Phase_C_P = round(Read_Phase_Power[8], 6)
    Phase_C_Q = round(Read_Phase_Power[9], 6)
    Phase_C_S = round(Read_Phase_Power[10], 6)
    Phase_C_PF = round(Read_Phase_Power[11], 6)

    if Power[8] != 0:
        scale_Phase_C_P = abs((Phase_C_P - Power[8]) / Power[8])
    else:
        if Power[8] == Phase_C_P == 0 or Phase_C_P < 0.001:
            scale_Phase_C_P = 0
        else:
            scale_Phase_C_P = 'null'

    if Power[9] != 0:
        scale_Phase_C_Q = abs((Phase_C_Q - Power[9]) / Power[9])
    else:
        if Power[9] == Phase_C_Q == 0 or Phase_C_Q < 0.001:
            scale_Phase_C_Q = 0
        else:
            scale_Phase_C_Q = 'null'

    if Power[10] != 0:
        scale_Phase_C_S = abs((Phase_C_S - Power[10]) / Power[10])
    else:
        if Power[10] == Phase_C_S == 0 or Phase_C_S < 0.001:
            scale_Phase_C_S = 0
        else:
            scale_Phase_C_S = 'null'

    if Power[11] != 0:
        scale_Phase_C_PF = abs((Phase_C_PF - Power[11]) / Power[11])
    else:
        if Power[11] == Phase_C_PF == 0 or Phase_C_PF < 0.001:
            scale_Phase_C_PF = 0
        else:
            scale_Phase_C_PF = 'null'

    Sys_P = round(Read_Phase_Power[12], 10)
    Sys_Q = round(Read_Phase_Power[13], 10)
    Sys_S = round(Read_Phase_Power[14], 10)
    Sys_PF = round(Read_Phase_Power[15], 10)

    if Power[12] != 0:
        scale_Sys_P = abs((Sys_P - Power[12]) / Power[12])
    else:
        if Power[12] == Sys_P == 0 or Sys_P < 0.001:
            scale_Sys_P = 0
        else:
            scale_Sys_P = 'null'

    if Power[13] != 0:
        scale_Sys_Q = abs((Sys_Q - Power[13]) / Power[13])
    else:
        if Power[13] == Sys_Q == 0 or Sys_Q < 0.001:
            scale_Sys_Q = 0
        else:
            scale_Sys_Q = 'null'

    if Power[14] != 0:
        scale_Sys_S = abs((Sys_S - Power[14]) / Power[14])
    else:
        if Power[14] == Sys_S == 0 or Sys_S < 0.001:
            scale_Sys_S = 0
        else:
            scale_Sys_S = 'null'

    if Power[15] != 0:
        scale_Sys_PF = abs((Sys_PF - Power[15]) / Power[15])
    else:
        if Power[15] == Sys_PF == 0 or Sys_PF < 0.001:
            scale_Sys_PF = 0
        else:
            scale_Sys_PF = 'null'

    if Power[12] != 0:
        scale_User1_P = abs((User1_P - Power[12]) / Power[12])
    else:
        if Power[12] == User1_P == 0 or User1_P < 0.001:
            scale_User1_P = 0
        else:
            scale_User1_P = 'null'

    if Power[13] != 0:
        scale_User1_Q = abs((User1_Q - Power[13]) / Power[13])
    else:
        if Power[13] == User1_Q == 0 or User1_Q < 0.001:
            scale_User1_Q = 0
        else:
            scale_User1_Q = 'null'

    if Power[14] != 0:
        scale_User1_S = abs((User1_S - Power[14]) / Power[14])
    else:
        if Power[14] == User1_S == 0 or User1_S < 0.001:
            scale_User1_S = 0
        else:
            scale_User1_S = 'null'

    if Power[15] != 0:
        scale_User1_PF = abs((User1_PF - Power[15]) / Power[15])
    else:
        if Power[15] == User1_PF == 0 or User1_PF < 0.001:
            scale_User1_PF = 0
        else:
            scale_User1_PF = 'null'

    power_list.extend(
        [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [User1_P, scale_User1_P], [User1_Q, scale_User1_Q],
         [User1_S, scale_User1_S], [User1_PF, scale_User1_PF]])
    return power_list


def e3_4w_y_pf_calculate_angle(PF):
    if PF == 1:
        return 120, 240, 0, 120, 240, 0
    if PF == '0.5L':
        return 120, 240, 0, 60, 180, 300
    if PF == '0.8C':
        return 120, 240, 0, 156.87, 276.87, 36.87


def e3_4w_y_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []
    Phase_A_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_S = round(apparent_power_calculate(Va, I1), 10)
    Phase_A_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    Phase_B_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_S = round(apparent_power_calculate(Vb, I2), 10)
    Phase_B_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    Phase_C_P = round(active_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    Phase_C_Q = round(reactive_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    Phase_C_S = round(apparent_power_calculate(Vc, I3), 10)
    Phase_C_PF = round(power_factor_calculate(Vc_ang, I3_ang), 10)

    Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
    Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
    Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend(
        [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
         Phase_C_Q,
         Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
    return Power


def e2_3w_network_pf_calculate_angle(PF):
    if PF == 1:
        return 120, 240, 0, 120, 240, 0
    if PF == '0.5L':
        return 120, 240, 0, 60, 180, 300
    if PF == '0.8C':
        return 120, 240, 0, 156.87, 276.87, 36.87


def e2_3w_network_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []
    Phase_A_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_S = round(apparent_power_calculate(Va, I1), 10)
    Phase_A_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    Phase_B_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_S = round(apparent_power_calculate(Vb, I2), 10)
    Phase_B_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    Phase_C_P = 0
    Phase_C_Q = 0
    Phase_C_S = 0
    Phase_C_PF = 0

    Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
    Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
    Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend(
        [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
         Phase_C_Q,
         Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
    return Power


def e3_3w_delta_pf_calculate_angle(PF):
    if PF == 1:
        return 120, 240, 0, 150, 270, 30
    if PF == '0.5L':
        return 120, 240, 0, 90, 210, 330
    if PF == '0.8C':
        return 120, 240, 0, 186.87, 306.87, 66.87


def e3_3w_delta_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []
    Va = Va * math.sqrt(3)
    Vb = Vb * math.sqrt(3)
    Vc = Vc * math.sqrt(3)
    Va_ang = Va_ang + 30
    Vb_ang = Vb_ang + 30
    Vc_ang = Vc_ang + 30
    Phase_A_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_S = round(apparent_power_calculate(Va, I1), 10)
    Phase_A_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    Phase_B_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    Phase_B_S = round(apparent_power_calculate(Vb, I2), 10)
    Phase_B_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    Phase_C_P = round(active_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    Phase_C_Q = round(reactive_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    Phase_C_S = round(apparent_power_calculate(Vc, I3), 10)
    Phase_C_PF = round(power_factor_calculate(Vc_ang, I3_ang), 10)

    Sys_P = Phase_A_P + Phase_B_P + Phase_C_P
    Sys_Q = Phase_A_Q + Phase_B_Q + Phase_C_Q
    Sys_S = Phase_A_S + Phase_B_S + Phase_C_S

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend(
        [Phase_A_P, Phase_A_Q, Phase_A_S, Phase_A_PF, Phase_B_P, Phase_B_Q, Phase_B_S, Phase_B_PF, Phase_C_P,
         Phase_C_Q,
         Phase_C_S, Phase_C_PF, Sys_P, Sys_Q, Sys_S, Sys_PF])
    return Power


def power_calculate(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang, Service):
    power_list = []
    Power = []
    if Service == "2 Element 3 Wire 1 Phase":
        Power = e2_3w_1p_power_standard_value(Va, Vc, I1, I2, Va_ang, I1_ang, Vc_ang, I2_ang)
    if Service == '3 Element 4 Wire Y':
        Power = e3_4w_y_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang)
    if Service == '2 Element 3 Wire Network':
        Power = e2_3w_network_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang,
                                                   I3_ang)
    if Service == '3 Element 3 Wire Delta':
        Power = e3_3w_delta_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang,
                                                 I3_ang)
    Read_Phase_Power = read_phase_sys_power(Power, times=10)
    User1_P = read_user_channel_1_active_power(Power[12], times=10)
    User1_Q = read_user_channel_1_reactive_power(Power[13], times=10)
    User1_S = read_user_channel_1_apparent_power(Power[14], times=10)
    User1_PF = read_user_channel_1_power_factor(Power[15], times=10)

    Phase_A_P = round(Read_Phase_Power[0], 6)
    Phase_A_Q = round(Read_Phase_Power[1], 6)
    Phase_A_S = round(Read_Phase_Power[2], 6)
    Phase_A_PF = round(Read_Phase_Power[3], 6)

    if Power[0] != 0:
        scale_Phase_A_P = abs((Phase_A_P - Power[0]) / Power[0])
    else:
        if Power[0] == Phase_A_P == 0 or Phase_A_P < 0.001:
            scale_Phase_A_P = 0
        else:
            scale_Phase_A_P = 'null'

    if Power[1] != 0:
        scale_Phase_A_Q = abs((Phase_A_Q - Power[1]) / Power[1])
    else:
        if Power[1] == Phase_A_Q == 0 or Phase_A_Q < 0.001:
            scale_Phase_A_Q = 0
        else:
            scale_Phase_A_Q = 'null'

    if Power[2] != 0:
        scale_Phase_A_S = abs((Phase_A_S - Power[2]) / Power[2])
    else:
        if Power[2] == Phase_A_S == 0 or Phase_A_S < 0.001:
            scale_Phase_A_S = 0
        else:
            scale_Phase_A_S = 'null'

    if Power[3] != 0:
        scale_Phase_A_PF = abs((Phase_A_PF - Power[3]) / Power[3])
    else:
        if Power[3] == Phase_A_PF == 0 or Phase_A_PF < 0.001:
            scale_Phase_A_PF = 0
        else:
            scale_Phase_A_PF = 'null'

    Phase_B_P = round(Read_Phase_Power[4], 6)
    Phase_B_Q = round(Read_Phase_Power[5], 6)
    Phase_B_S = round(Read_Phase_Power[6], 6)
    Phase_B_PF = round(Read_Phase_Power[7], 6)

    if Power[4] != 0:
        scale_Phase_B_P = abs((Phase_B_P - Power[4]) / Power[4])
    else:
        if Power[4] == Phase_B_P == 0 or Phase_B_P < 0.001:
            scale_Phase_B_P = 0
        else:
            scale_Phase_B_P = 'null'

    if Power[5] != 0:
        scale_Phase_B_Q = abs((Phase_B_Q - Power[5]) / Power[5])
    else:
        if Power[5] == Phase_B_Q == 0 or Phase_B_Q < 0.001:
            scale_Phase_B_Q = 0
        else:
            scale_Phase_B_Q = 'null'

    if Power[6] != 0:
        scale_Phase_B_S = abs((Phase_B_S - Power[6]) / Power[6])
    else:
        if Power[6] == Phase_B_S == 0 or Phase_B_S < 0.001:
            scale_Phase_B_S = 0
        else:
            scale_Phase_B_S = 'null'

    if Power[7] != 0:
        scale_Phase_B_PF = abs((Phase_B_PF - Power[7]) / Power[7])
    else:
        if Power[7] == Phase_B_PF == 0 or Phase_B_PF < 0.001:
            scale_Phase_B_PF = 0
        else:
            scale_Phase_B_PF = 'null'

    Phase_C_P = round(Read_Phase_Power[8], 6)
    Phase_C_Q = round(Read_Phase_Power[9], 6)
    Phase_C_S = round(Read_Phase_Power[10], 6)
    Phase_C_PF = round(Read_Phase_Power[11], 6)

    if Power[8] != 0:
        scale_Phase_C_P = abs((Phase_C_P - Power[8]) / Power[8])
    else:
        if Power[8] == Phase_C_P == 0 or Phase_C_P < 0.001:
            scale_Phase_C_P = 0
        else:
            scale_Phase_C_P = 'null'

    if Power[9] != 0:
        scale_Phase_C_Q = abs((Phase_C_Q - Power[9]) / Power[9])
    else:
        if Power[9] == Phase_C_Q == 0 or Phase_C_Q < 0.001:
            scale_Phase_C_Q = 0
        else:
            scale_Phase_C_Q = 'null'

    if Power[10] != 0:
        scale_Phase_C_S = abs((Phase_C_S - Power[10]) / Power[10])
    else:
        if Power[10] == Phase_C_S == 0 or Phase_C_S < 0.001:
            scale_Phase_C_S = 0
        else:
            scale_Phase_C_S = 'null'

    if Power[11] != 0:
        scale_Phase_C_PF = abs((Phase_C_PF - Power[11]) / Power[11])
    else:
        if Power[11] == Phase_C_PF == 0 or Phase_C_PF < 0.001:
            scale_Phase_C_PF = 0
        else:
            scale_Phase_C_PF = 'null'

    Sys_P = round(Read_Phase_Power[12], 10)
    Sys_Q = round(Read_Phase_Power[13], 10)
    Sys_S = round(Read_Phase_Power[14], 10)
    Sys_PF = round(Read_Phase_Power[15], 10)

    if Power[12] != 0:
        scale_Sys_P = abs((Sys_P - Power[12]) / Power[12])
    else:
        if Power[12] == Sys_P == 0 or Sys_P < 0.001:
            scale_Sys_P = 0
        else:
            scale_Sys_P = 'null'

    if Power[13] != 0:
        scale_Sys_Q = abs((Sys_Q - Power[13]) / Power[13])
    else:
        if Power[13] == Sys_Q == 0 or Sys_Q < 0.001:
            scale_Sys_Q = 0
        else:
            scale_Sys_Q = 'null'

    if Power[14] != 0:
        scale_Sys_S = abs((Sys_S - Power[14]) / Power[14])
    else:
        if Power[14] == Sys_S == 0 or Sys_S < 0.001:
            scale_Sys_S = 0
        else:
            scale_Sys_S = 'null'

    if Power[15] != 0:
        scale_Sys_PF = abs((Sys_PF - Power[15]) / Power[15])
    else:
        if Power[15] == Sys_PF == 0 or Sys_PF < 0.001:
            scale_Sys_PF = 0
        else:
            scale_Sys_PF = 'null'

    if Power[12] != 0:
        scale_User1_P = abs((User1_P - Power[12]) / Power[12])
    else:
        if Power[12] == User1_P == 0 or User1_P < 0.001:
            scale_User1_P = 0
        else:
            scale_User1_P = 'null'

    if Power[13] != 0:
        scale_User1_Q = abs((User1_Q - Power[13]) / Power[13])
    else:
        if Power[13] == User1_Q == 0 or User1_Q < 0.001:
            scale_User1_Q = 0
        else:
            scale_User1_Q = 'null'

    if Power[14] != 0:
        scale_User1_S = abs((User1_S - Power[14]) / Power[14])
    else:
        if Power[14] == User1_S == 0 or User1_S < 0.001:
            scale_User1_S = 0
        else:
            scale_User1_S = 'null'

    if Power[15] != 0:
        scale_User1_PF = abs((User1_PF - Power[15]) / Power[15])
    else:
        if Power[15] == User1_PF == 0 or User1_PF < 0.001:
            scale_User1_PF = 0
        else:
            scale_User1_PF = 'null'

    power_list.extend(
        [[Phase_A_P, scale_Phase_A_P], [Phase_A_Q, scale_Phase_A_Q], [Phase_A_S, scale_Phase_A_S],
         [Phase_A_PF, scale_Phase_A_PF], [Phase_B_P, scale_Phase_B_P], [Phase_B_Q, scale_Phase_B_Q],
         [Phase_B_S, scale_Phase_B_S], [Phase_B_PF, scale_Phase_B_PF], [Phase_C_P, scale_Phase_C_P],
         [Phase_C_Q, scale_Phase_C_Q],
         [Phase_C_S, scale_Phase_C_S], [Phase_C_PF, scale_Phase_C_PF], [Sys_P, scale_Sys_P], [Sys_Q, scale_Sys_Q],
         [Sys_S, scale_Sys_S], [Sys_PF, scale_Sys_PF], [User1_P, scale_User1_P], [User1_Q, scale_User1_Q],
         [User1_S, scale_User1_S], [User1_PF, scale_User1_PF]])
    return power_list


def e1_2w_pf_calculate_angle(PF):
    if PF == 1:
        return 0, 0, 0, 0, 0, 0
    if PF == 0.5:
        return 0, 0, 0, 0, 0, 300
    if PF == -0.8:
        return 0, 0, 0, 0, 0, 216.87


def read_phase_a_voltage_new(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr['Phase A Line-to-Neutral Voltage']['Start(Dec)'],
                                              count=real_time_addr['Phase A Line-to-Neutral Voltage']['Reg'], slave=1)
        logging.info('Phase_A_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    avg_val = sum(val_list) / len(val_list)
    avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
    min_val_accuracy = round(abs((val_list[0] - standard_value) / standard_value), 5)
    max_val_accuracy = round(abs((val_list[-1] - standard_value) / standard_value), 5)
    return [[val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]]


def read_phase_a_voltage_angle_new(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=real_time_addr[
            'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
            'Start(Dec)'],
                                              count=real_time_addr[
                                                  'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
                                                  'Reg'], slave=1)
        logging.info('Phase_A_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def read_input_current(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Current']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input_Channel_{input_ch}_Current ret is:{value}')
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
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
    return [[val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]]


def read_input_active_power(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Active Power']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Active Power']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Channel_{input_ch}_Active_Power ret is:{value}')
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
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
    return [[val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]]


def read_input_channel_reactive_power_new(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Reactive Power']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Reactive Power']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input_Channel_{input_ch}_Reactive_Power ret is:{value}')
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
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
    return [[val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]]


def read_input_channel_apparent_power(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Apparent Power']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Apparent Power']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input_Channel_{input_ch}_Apparent_Power ret is:{value}')
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
            avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
    return [[val_list[0], min_val_accuracy], [val_list[-1], max_val_accuracy], [avg_val, avg_val_accuracy]]


def read_input_channel_current_angle_new(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current Phase Angle']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Current Phase Angle']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input_Channel_{input_ch}_Current_Phase_Angle ret is:{value}')
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    if standard_value == 0:
        for j in range(len(val_list)):
            if 350 <= val_list[j] <= 360:
                val_list[j] = val_list[j] - 360
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        if val_list[-1] < 0:
            val_list[-1] = val_list[-1] + 360
        return val_list[-1]
    if val_list[0] < 0:
        val_list[0] = val_list[0] + 360
    return val_list[0]


def read_phase_l_n_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Phase A Line-to-Neutral Voltage']['Start(Dec)']
    count = real_time_addr['Phase A Line-to-Neutral Voltage']['Reg'] * 3
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('read_phase_l_n_voltage ret is:{}'.format(value))
        read_list = []
        for i in range(3):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(3):
        Phase_L_N_Voltage = [x[i] for x in val_list]
        min_Phase_L_N_Voltage = min(Phase_L_N_Voltage)
        max_Phase_L_N_Voltage = max(Phase_L_N_Voltage)
        avg_Phase_L_N_Voltage = sum(Phase_L_N_Voltage) / len(Phase_L_N_Voltage)
        if standard_value[i] != 0:
            min_val_accuracy = round(abs((min_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            max_val_accuracy = round(abs((max_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            avg_val_accuracy = round(abs((avg_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
        else:
            if avg_Phase_L_N_Voltage == standard_value[
                i] == min_Phase_L_N_Voltage == max_Phase_L_N_Voltage == 0:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
            else:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
        ret_list.extend([[min_Phase_L_N_Voltage, min_val_accuracy], [max_Phase_L_N_Voltage, max_val_accuracy],
                         [avg_Phase_L_N_Voltage, avg_val_accuracy]])
    return ret_list


def read_phase_angle_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr[
        'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
        'Start(Dec)']
    count = real_time_addr[
                'Phase A Line-to-Neutral Voltage Phase Angle(1E_2W、2E_3W_1P、3E_4W_Y、2E_3W_Network)  /\nPhase AB Line-to-Linel Voltage Phase Angle (2E_3W_Delta、 3E_3W_Delta)'][
                'Reg'] * 3
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('read_phase_l_n_voltage ret is:{}'.format(value))
        read_list = []
        for i in range(3):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(3):
        Phase_L_N_Voltage = [x[i] for x in val_list]
        if standard_value[i] == 0:
            # filtered_data = [subarray for subarray in Phase_L_N_Voltage if 350 <= subarray[i] <= 360]
            for j in range(len(Phase_L_N_Voltage)):
                if 350 <= Phase_L_N_Voltage[j] <= 360:
                    Phase_L_N_Voltage[j] = Phase_L_N_Voltage[j] - 360
        min_Phase_L_N_Voltage = min(Phase_L_N_Voltage)
        max_Phase_L_N_Voltage = max(Phase_L_N_Voltage)
        avg_Phase_L_N_Voltage = sum(Phase_L_N_Voltage) / len(Phase_L_N_Voltage)
        min_val_accuracy = round(abs(min_Phase_L_N_Voltage - standard_value[i]), 5)
        max_val_accuracy = round(abs(max_Phase_L_N_Voltage - standard_value[i]), 5)
        avg_val_accuracy = round(abs(avg_Phase_L_N_Voltage - standard_value[i]), 5)
        if min_Phase_L_N_Voltage < 0:
            min_Phase_L_N_Voltage = min_Phase_L_N_Voltage + 360
        if max_Phase_L_N_Voltage < 0:
            max_Phase_L_N_Voltage = max_Phase_L_N_Voltage + 360
        if avg_Phase_L_N_Voltage < 0:
            avg_Phase_L_N_Voltage = avg_Phase_L_N_Voltage + 360
        ret_list.extend([[min_Phase_L_N_Voltage, min_val_accuracy], [max_Phase_L_N_Voltage, max_val_accuracy],
                         [avg_Phase_L_N_Voltage, avg_val_accuracy]])
    return ret_list


def read_system_power(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['System Active Power']['Start(Dec)']
    count = real_time_addr['System Active Power']['Reg'] * 3
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('System P/Q/S Power ret is:{}'.format(value))
        read_list = []
        for i in range(3):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(3):
        Phase_L_N_Voltage = [x[i] for x in val_list]
        min_Phase_L_N_Voltage = min(Phase_L_N_Voltage)
        max_Phase_L_N_Voltage = max(Phase_L_N_Voltage)
        avg_Phase_L_N_Voltage = sum(Phase_L_N_Voltage) / len(Phase_L_N_Voltage)
        if standard_value[i] != 0:
            min_val_accuracy = round(abs((min_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            max_val_accuracy = round(abs((max_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            avg_val_accuracy = round(abs((avg_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
        else:
            if avg_Phase_L_N_Voltage == standard_value[
                i] == min_Phase_L_N_Voltage == max_Phase_L_N_Voltage == 0:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
            else:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
        ret_list.extend([[min_Phase_L_N_Voltage, min_val_accuracy], [max_Phase_L_N_Voltage, max_val_accuracy],
                         [avg_Phase_L_N_Voltage, avg_val_accuracy]])
    return ret_list


def read_input_power(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Current']['Reg'] * 4
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input Channel {input_ch} Current/P/Q/S ret is:{value}')
        read_list = []
        for i in range(4):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(4):
        Phase_L_N_Voltage = [x[i] for x in val_list]
        min_Phase_L_N_Voltage = min(Phase_L_N_Voltage)
        max_Phase_L_N_Voltage = max(Phase_L_N_Voltage)
        avg_Phase_L_N_Voltage = sum(Phase_L_N_Voltage) / len(Phase_L_N_Voltage)
        if standard_value[i] != 0:
            min_val_accuracy = round(abs((min_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            max_val_accuracy = round(abs((max_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            avg_val_accuracy = round(abs((avg_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
        else:
            if avg_Phase_L_N_Voltage == standard_value[
                i] == min_Phase_L_N_Voltage == max_Phase_L_N_Voltage == 0:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
            else:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
        ret_list.extend([[min_Phase_L_N_Voltage, min_val_accuracy], [max_Phase_L_N_Voltage, max_val_accuracy],
                         [avg_Phase_L_N_Voltage, avg_val_accuracy]])
    return ret_list


def read_user_power_nwe(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['User Channel 1 Active Power']['Start(Dec)']
    count = real_time_addr['User Channel 1 Active Power']['Reg'] * 3
    val_list = []
    address = (input_ch - 1) * 12 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'User Channel {input_ch} P/Q/S Power ret is:{value}')
        read_list = []
        for i in range(3):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(3):
        value_list = [x[i] for x in val_list]
        min_value_list = min(value_list)
        max_value_list = max(value_list)
        avg_value_list = sum(value_list) / len(value_list)
        if standard_value[i] != 0:
            min_val_accuracy = round(abs((min_value_list - standard_value[i]) / standard_value[i]), 5)
            max_val_accuracy = round(abs((max_value_list - standard_value[i]) / standard_value[i]), 5)
            avg_val_accuracy = round(abs((avg_value_list - standard_value[i]) / standard_value[i]), 5)
        else:
            if avg_value_list == standard_value[
                i] == min_value_list == max_value_list == 0:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
            else:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 1
        ret_list.extend([[min_value_list, min_val_accuracy], [max_value_list, max_val_accuracy],
                         [avg_value_list, avg_val_accuracy]])
    return ret_list


def read_input_angle_current(input_ch, standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Input Channel 1 Current Phase Angle']['Start(Dec)']
    count = real_time_addr['Input Channel 1 Current Phase Angle']['Reg']
    val_list = []
    address = (input_ch - 1) * 14 + address
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info(f'Input Channel {input_ch} Current Phase Angle ret is:{value}')
        read_list = []
        for i in range(1):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(1):
        value_list = [x[i] for x in val_list]
        if standard_value[i] == 0:
            for j in range(len(value_list)):
                if 350 <= value_list[j] <= 360:
                    value_list[j] = value_list[j] - 360
        min_value_list = min(value_list)
        max_value_list = max(value_list)
        avg_value_list = sum(value_list) / len(value_list)
        min_val_accuracy = round(abs(min_value_list - standard_value[i]), 5)
        max_val_accuracy = round(abs(max_value_list - standard_value[i]), 5)
        avg_val_accuracy = round(abs(avg_value_list - standard_value[i]), 5)
        if min_value_list < 0:
            min_value_list = min_value_list + 360
        if max_value_list < 0:
            max_value_list = max_value_list + 360
        if avg_value_list < 0:
            avg_value_list = avg_value_list + 360
        ret_list.extend([[min_value_list, min_val_accuracy], [max_value_list, max_val_accuracy],
                         [avg_value_list, avg_val_accuracy]])
    return ret_list


def e3_4w_y_pf_calculate_angle_new(PF):
    if PF == 1:
        return 120, 240, 0, 120, 240, 0
    if PF == 0.5:
        return 120, 240, 0, 60, 180, 300
    if PF == -0.8:
        return 120, 240, 0, 336.87, 96.87, 216.87


def e3_4w_y_input_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang,
                                       input_ch_num):
    Power = []
    input_1_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_S = round(apparent_power_calculate(Va, I1), 10)
    input_1_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    input_2_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_S = round(apparent_power_calculate(Vb, I2), 10)
    input_2_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    input_3_P = round(active_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_Q = round(reactive_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_S = round(apparent_power_calculate(Vc, I3), 10)
    input_3_PF = round(power_factor_calculate(Vc_ang, I3_ang), 10)

    input_ch = input_ch_num / 3
    Sys_P = (input_1_P + input_2_P + input_3_P) * input_ch
    Sys_Q = (input_1_Q + input_2_Q + input_3_Q) * input_ch
    Sys_S = (input_1_S + input_2_S + input_3_S) * input_ch

    User_P = (input_1_P + input_2_P + input_3_P)
    User_Q = (input_1_Q + input_2_Q + input_3_Q)
    User_S = (input_1_S + input_2_S + input_3_S)
    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend([[I1, input_1_P, input_1_Q, input_1_S], [I2, input_2_P, input_2_Q, input_2_S],
                  [I3, input_3_P, input_3_Q, input_3_S], [Sys_P, Sys_Q, Sys_S], [User_P, User_Q, User_S]])
    return Power


def e2_3w_network_input_power_standard_value(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []
    input_1_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_S = round(apparent_power_calculate(Va, I1), 10)
    input_1_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    input_2_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_S = round(apparent_power_calculate(Vb, I2), 10)
    input_2_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    input_3_P = round(active_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_Q = round(reactive_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_S = round(apparent_power_calculate(Vc, I3), 10)
    input_3_PF = round(power_factor_calculate(Vc_ang, I3_ang), 10)

    Sys_P = (input_1_P + input_2_P + input_3_P) * 8
    Sys_Q = (input_1_Q + input_2_Q + input_3_Q) * 8
    Sys_S = (input_1_S + input_2_S + input_3_S) * 8

    User_P = (input_1_P + input_2_P)
    User_Q = (input_1_Q + input_2_Q)
    User_S = (input_1_S + input_2_S)

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend([[I1, input_1_P, input_1_Q, input_1_S], [I2, input_2_P, input_2_Q, input_2_S],
                  [I3, input_3_P, input_3_Q, input_3_S], [Sys_P, Sys_Q, Sys_S], [User_P, User_Q, User_S]])
    return Power


def e2_3w_1p_pf_calculate_angle_new(PF):
    if PF == 1:
        return 180, 0, 0, 0, 180, 0
    if PF == 0.5:
        return 180, 0, 0, 0, 120, 300
    if PF == -0.8:
        return 180, 0, 0, 0, 36.87, 216.87


def e2_3w_1p_input_power_standard_value(Va, Vc, I1, I2, Va_ang, I1_ang, Vc_ang, I2_ang):
    Power = []
    Phase_A_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    Phase_A_S = round(apparent_power_calculate(Va, I1), 10)
    Phase_A_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    Phase_B_P = 0
    Phase_B_Q = 0
    Phase_B_S = 0
    Phase_B_PF = 1

    Phase_C_P = round(active_power_calculate(Vc, I2, Vc_ang, I2_ang), 10)
    Phase_C_Q = round(reactive_power_calculate(Vc, I2, Vc_ang, I2_ang), 10)
    Phase_C_S = round(apparent_power_calculate(Vc, I2), 10)
    Phase_C_PF = round(power_factor_calculate(Vc_ang, I2_ang), 10)

    Sys_P = (Phase_A_P + Phase_B_P + Phase_C_P) * 12
    Sys_Q = (Phase_A_Q + Phase_B_Q + Phase_C_Q) * 12
    Sys_S = (Phase_A_S + Phase_B_S + Phase_C_S) * 12

    User_P = (Phase_A_P + Phase_B_P + Phase_C_P)
    User_Q = (Phase_A_Q + Phase_B_Q + Phase_C_Q)
    User_S = (Phase_A_S + Phase_B_S + Phase_C_S)

    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend([[I1, Phase_A_P, Phase_A_Q, Phase_A_S], [I2, Phase_C_P, Phase_C_Q, Phase_C_S],
                  [0, Phase_B_P, Phase_B_Q, Phase_B_S],
                  [Sys_P, Sys_Q, Sys_S], [User_P, User_Q, User_S]])
    return Power


def e3_3w_delta_pf_calculate_angle_new(PF):
    if PF == 1:
        return 120, 240, 0, 150, 270, 30
    if PF == 0.5:
        return 120, 240, 0, 90, 210, 330
    if PF == -0.8:
        return 120, 240, 0, 6.87, 126.87, 246.87


def read_phase_l_l_voltage(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Phase AB Line-to-Line Voltage']['Start(Dec)']
    count = real_time_addr['Phase AB Line-to-Line Voltage']['Reg'] * 3
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('read_phase_l_l_voltage ret is:{}'.format(value))
        read_list = []
        for i in range(3):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(3):
        Phase_L_N_Voltage = [x[i] for x in val_list]
        min_Phase_L_N_Voltage = min(Phase_L_N_Voltage)
        max_Phase_L_N_Voltage = max(Phase_L_N_Voltage)
        avg_Phase_L_N_Voltage = sum(Phase_L_N_Voltage) / len(Phase_L_N_Voltage)
        if standard_value[i] != 0:
            min_val_accuracy = round(abs((min_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            max_val_accuracy = round(abs((max_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
            avg_val_accuracy = round(abs((avg_Phase_L_N_Voltage - standard_value[i]) / standard_value[i]), 5)
        else:
            if avg_Phase_L_N_Voltage == standard_value[
                i] == min_Phase_L_N_Voltage == max_Phase_L_N_Voltage == 0:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 0
            else:
                avg_val_accuracy = min_val_accuracy = max_val_accuracy = 'null'
        ret_list.extend([[min_Phase_L_N_Voltage, min_val_accuracy], [max_Phase_L_N_Voltage, max_val_accuracy],
                         [avg_Phase_L_N_Voltage, avg_val_accuracy]])
    return ret_list


def e3_3w_delta_power_standard_value_new(Va, Vb, Vc, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []
    I1_ang = I1_ang - 30
    if I1_ang < 0:
        I1_ang = I1_ang + 360
    I2_ang = I2_ang - 30
    if I2_ang < 0:
        I2_ang = I2_ang + 360
    I3_ang = I3_ang - 30
    if I3_ang < 0:
        I3_ang = I3_ang + 360
    input_1_P = round(active_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_Q = round(reactive_power_calculate(Va, I1, Va_ang, I1_ang), 10)
    input_1_S = round(apparent_power_calculate(Va, I1), 10)
    input_1_PF = round(power_factor_calculate(Va_ang, I1_ang), 10)

    input_2_P = round(active_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_Q = round(reactive_power_calculate(Vb, I2, Vb_ang, I2_ang), 10)
    input_2_S = round(apparent_power_calculate(Vb, I2), 10)
    input_2_PF = round(power_factor_calculate(Vb_ang, I2_ang), 10)

    input_3_P = round(active_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_Q = round(reactive_power_calculate(Vc, I3, Vc_ang, I3_ang), 10)
    input_3_S = round(apparent_power_calculate(Vc, I3), 10)
    input_3_PF = round(power_factor_calculate(Vc_ang, I3_ang), 10)

    Sys_P = (input_1_P + input_2_P + input_3_P) * 8
    Sys_Q = (input_1_Q + input_2_Q + input_3_Q) * 8
    Sys_S = (input_1_S + input_2_S + input_3_S) * 8

    User_P = (input_1_P + input_2_P + input_3_P)
    User_Q = (input_1_Q + input_2_Q + input_3_Q)
    User_S = (input_1_S + input_2_S + input_3_S)
    if Sys_S != 0:
        Sys_PF = round(Sys_P / Sys_S, 3)
    else:
        Sys_PF = 0
    Power.extend([[I1, input_1_P, input_1_Q, input_1_S], [I2, input_2_P, input_2_Q, input_2_S],
                  [I3, input_3_P, input_3_Q, input_3_S], [Sys_P, Sys_Q, Sys_S], [User_P, User_Q, User_S]])
    return Power


def e2_3w_delta_pf_calculate_angle_new(PF):
    if PF == 1:
        return 120, 240, 0, 0, 120, 0
    if PF == 0.5:
        return 120, 240, 0, 0, 60, 300
    if PF == -0.8:
        return 120, 240, 0, 0, 336.87, 216.87


def e2_3w_delta_power_standard_value_new(Vab, Vbc, Vca, I1, I2, I3, Va_ang, I1_ang, Vb_ang, I2_ang, Vc_ang, I3_ang):
    Power = []

    input_1_P = 0
    input_1_Q = 0
    input_1_S = 0

    input_2_P = 0
    input_2_Q = 0
    input_2_S = 0

    input_3_P = 0
    input_3_Q = 0
    input_3_S = 0

    User_S = ((Vab * I1 + Vbc * I2) * math.cos(math.radians(30))) / 1000
    User_P = User_S * (math.cos(math.radians(Va_ang - I1_ang)))
    User_Q = math.sqrt((User_S ** 2) - (User_P ** 2))
    if math.sin(math.radians(Va_ang - I1_ang)) >= 0:
        User_Q = User_Q
    else:
        User_Q = -User_Q
    Sys_P = User_P * 12
    Sys_Q = User_Q * 12
    Sys_S = User_S * 12

    Power.extend([[I1, input_1_P, input_1_Q, input_1_S], [I2, input_2_P, input_2_Q, input_2_S],
                  [I3, input_3_P, input_3_Q, input_3_S], [Sys_P, Sys_Q, Sys_S], [User_P, User_Q, User_S]])
    return Power


def read_phase_sys_power_111(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = real_time_addr['Phase A Current']['Start(Dec)']
    count = real_time_addr['Phase A Current']['Reg']
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=48, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
        value = [value[i] for i in range(len(value)) if
                 i not in {0, 1, 8, 9, 12, 13, 20, 21, 24, 25, 32, 33, 36, 37, 44, 45}]
        read_list = []
        for i in range(16):
            i = i * 2
            reg = hex(value[i]).replace('0x', '').zfill(4) + hex(value[i + 1]).replace('0x', '').zfill(4)
            hex_num = reg.replace('0x', '')
            integer_num = int(hex_num, 16)
            value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
            read_list.append(value_measu)
        val_list.append(read_list)
    ret_list = []
    for i in range(16):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list


def read_baud_rate(standard_value, times=1):
    # ModbusClient = modbusrtuortcp(conn_mode=modbus_config['conn_mode'])
    address = basic_setting['Baud Rate']['Start(Dec)']
    count = basic_setting['Baud Rate']['Reg']
    value = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=address, count=count, slave=1)
        logging.info('Phase_B_Active_Power ret is:{}'.format(value))
    return value[0]
