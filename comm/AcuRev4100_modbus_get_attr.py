import logging
from modbus_config import modbus_config
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from tools.excel_operate import dcpara_4100addr_get
import struct


def read_system_frequency(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=8192, count=2, slave=1)
    client.close()
    reg = hex(voltage[0]).replace('0x', '').zfill(4) + hex(voltage[1]).replace('0x', '').zfill(4)
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_frequency(standard_value, times=1):
    mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        value = mes.read_measurement(address=8192, count=2, slave=1)
        logging.info('frequency ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    mes.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def Read_Phase_A_Voltage(standard_value, times=1):
    mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        value = mes.read_measurement(address=8194, count=2, slave=1)
        logging.info('Phase_A_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    mes.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def Read_Phase_B_Voltage(standard_value, times=1):
    mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        value = mes.read_measurement(address=8196, count=2, slave=1)
        logging.info('Phase_B_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    mes.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]


def Read_Phase_C_Voltage(standard_value, times=1):
    mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        value = mes.read_measurement(address=8198, count=2, slave=1)
        logging.info('Phase_C_Voltage ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    mes.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        return val_list[-1]
    return val_list[0]

