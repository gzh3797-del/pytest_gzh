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

# volt_cur_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])

ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def read_frequency(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8192, count=2, slave=1)
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


def Read_Phase_A_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8194, count=2, slave=1)
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


def Read_Phase_B_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8196, count=2, slave=1)
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

print(Read_Phase_B_Voltage(69, times=20))
def Read_Phase_C_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8198, count=2, slave=1)
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


def Read_Average_ln_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8200, count=2, slave=1)
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


def Read_Phase_AB_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8202, count=2, slave=1)
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


def Read_Phase_BC_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8204, count=2, slave=1)
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


def Read_Phase_CA_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8206, count=2, slave=1)
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


def Read_Average_ll_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8208, count=2, slave=1)
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


def Read_Phase_A_Voltage_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8210, count=2, slave=1)
        logging.info('Phase_A_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        # return val_list[-1]
        if abs(val_list[-1]-360) >= 359.9 and standard_value == 360:
            val_list[-1] = 360
            return val_list[-1]
        elif abs(val_list[-1]-360) >= 359.9 and standard_value == 0:
            val_list[-1] = 0
            return val_list[-1]
        else:
            return val_list[-1]

    if abs(val_list[0]-360) >= 359.9 and standard_value == 360:
        val_list[0] = 360
        return val_list[0]
    elif abs(val_list[0] - 360) >= 359.9 and standard_value == 0:
        val_list[0] = 0
        return val_list[0]
    else:
        return val_list[0]


def Read_Phase_B_Voltage_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8212, count=2, slave=1)
        logging.info('Phase_B_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        if abs(val_list[-1] - 360) >= 359.9:
            val_list[-1] = 360
            return val_list[-1]
        else:
            return val_list[-1]

    if abs(val_list[0] - 360) >= 359.9:
        val_list[0] = 360
        return val_list[0]
    else:
        return val_list[0]


def Read_Phase_C_Voltage_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8214, count=2, slave=1)
        logging.info('Phase_C_Voltage_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        if abs(val_list[-1] - 360) >= 359.9:
            val_list[-1] = 360
            return val_list[-1]
        else:
            return val_list[-1]

    if abs(val_list[0] - 360) >= 359.9:
        val_list[0] = 360
        return val_list[0]
    else:
        return val_list[0]


def Read_Phase_A_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8216, count=2, slave=1)
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


def Read_Phase_A_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
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
    print(val_list)
    ret_list = []
    for i in range(4):
        val_list.sort(key=lambda x: x[i])
        if abs(val_list[-1][i] - standard_value[i]) > abs(val_list[0][i] - standard_value[i]):
            ret_list.append(val_list[-1][i])
        else:
            ret_list.append(val_list[0][i])
    return ret_list
    # ModbusClient.close()


# print(Read_Phase_A_Power([0,0,0,0],5))


def Read_Phase_A_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8218, count=2, slave=1)
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


def Read_Phase_A_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8220, count=2, slave=1)
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


def Read_Phase_A_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8222, count=2, slave=1)
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


def Read_Phase_A_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8224, count=2, slave=1)
    logging.info('Phase_A_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def Read_Phase_A_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8226, count=2, slave=1)
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


def Read_Phase_B_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8228, count=2, slave=1)
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


def Read_Phase_B_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8230, count=2, slave=1)
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


def Read_Phase_B_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8232, count=2, slave=1)
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


def Read_Phase_B_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8234, count=2, slave=1)
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


def Read_Phase_B_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8236, count=2, slave=1)
    logging.info('Phase_B_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def Read_Phase_B_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8238, count=2, slave=1)
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


def Read_Phase_C_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8240, count=2, slave=1)
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


def Read_Phase_C_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8242, count=2, slave=1)
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


def Read_Phase_C_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8244, count=2, slave=1)
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


def Read_Phase_C_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8246, count=2, slave=1)
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


def Read_Phase_C_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8248, count=2, slave=1)
    logging.info('Phase_C_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def Read_Phase_C_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8250, count=2, slave=1)
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


def Read_System_Average_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8252, count=2, slave=1)
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


def Read_System_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8254, count=2, slave=1)
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


def Read_System_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8256, count=2, slave=1)
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


def Read_System_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8258, count=2, slave=1)
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


def Read_System_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8260, count=2, slave=1)
    logging.info('System_Load_Nature ret is:{}'.format(value))
    Load_Nature = ''
    if value[1] == 0:
        Load_Nature = 'R'
    if value[1] == 1:
        Load_Nature = 'C'
    if value[1] == 2:
        Load_Nature = 'L'
    return Load_Nature


def Read_System_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8262, count=2, slave=1)
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


def Read_Input_Channel_1_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8264, count=2, slave=1)
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


def Read_Input_Channel_1_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8266, count=2, slave=1)
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


def Read_Input_Channel_1_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8268, count=2, slave=1)
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


def Read_Input_Channel_1_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8270, count=2, slave=1)
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


def Read_Input_Channel_1_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8272, count=2, slave=1)
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


def Read_Input_Channel_1_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8274, count=2, slave=1)
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


def Read_Input_Channel_1_Current_Phase_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8276, count=2, slave=1)
        logging.info('Input_Channel_1_Current_Phase_Angle ret is:{}'.format(value))
        reg = hex(value[0]).replace('0x', '').zfill(4) + hex(value[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        val_list.append(value_measu)
    # ModbusClient.close()
    val_list.sort()
    if abs(val_list[-1] - standard_value) > abs(val_list[0] - standard_value):
        # return val_list[-1]
        if standard_value == 0:
            val_list[-1] = 0
            return val_list[-1]
        elif abs(val_list[-1] - 360) <=0.1 and standard_value ==360:
            val_list[-1] = 360
            return val_list[-1]
        else:
            return val_list[-1]
    # return val_list[0]
    if val_list[0] <= 0.1 and standard_value == 0:
        val_list[0] = 0
        return val_list[0]
    elif val_list[0] <= 0.1 and standard_value == 360:
        val_list[0] = 360
        return val_list[0]
    elif abs(val_list[0] - 360) <= 0.1 and standard_value ==360:
        val_list[0] = 360
        return val_list[0]
    else:
        return val_list[0]


def Read_Input_Channel_2_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8278, count=2, slave=1)
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


def Read_Input_Channel_2_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8280, count=2, slave=1)
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


def Read_Input_Channel_2_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8282, count=2, slave=1)
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


def Read_Input_Channel_2_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8284, count=2, slave=1)
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


def Read_Input_Channel_2_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8286, count=2, slave=1)
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


def Read_Input_Channel_2_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8288, count=2, slave=1)
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


def Read_Input_Channel_2_Current_Phase_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8290, count=2, slave=1)
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


def Read_Input_Channel_3_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8292, count=2, slave=1)
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


def Read_Input_Channel_3_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8294, count=2, slave=1)
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


def Read_Input_Channel_3_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8296, count=2, slave=1)
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


def Read_Input_Channel_3_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8298, count=2, slave=1)
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


def Read_Input_Channel_3_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8300, count=2, slave=1)
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


def Read_Input_Channel_3_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8302, count=2, slave=1)
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


def Read_Input_Channel_3_Current_Phase_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8304, count=2, slave=1)
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


def Read_User_Channel_1_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8600, count=2, slave=1)
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


def Read_User_Channel_1_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8602, count=2, slave=1)
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


def Read_User_Channel_1_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8604, count=2, slave=1)
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


def Read_User_Channel_1_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8606, count=2, slave=1)
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


def Read_User_Channel_1_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8608, count=2, slave=1)
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


def Read_User_Channel_1_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8610, count=2, slave=1)
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


def Read_User_Channel_2_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8612, count=2, slave=1)
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


def Read_User_Channel_2_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8614, count=2, slave=1)
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


def Read_User_Channel_2_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8616, count=2, slave=1)
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


def Read_User_Channel_2_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8618, count=2, slave=1)
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


def Read_User_Channel_2_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8620, count=2, slave=1)
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


def Read_User_Channel_2_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8622, count=2, slave=1)
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


def Read_User_Channel_3_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8624, count=2, slave=1)
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


def Read_User_Channel_3_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8626, count=2, slave=1)
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


def Read_User_Channel_3_Reactive_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8628, count=2, slave=1)
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


def Read_User_Channel_3_Apparent_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8630, count=2, slave=1)
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


def Read_User_Channel_3_Load_Nature():
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    value = ModbusClient.read_measurement(address=8632, count=2, slave=1)
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


def Read_User_Channel_3_Power_Factor(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8634, count=2, slave=1)
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


def Active_Power_calculate(voltage, current, voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Active_Power = (voltage * current * math.cos(math.radians(voltage_current_angle))) / 1000
    Active_Power = Active_Power
    return Active_Power


def Reactive_Power_calculate(voltage, current, voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Reactive_Power = (voltage * current * math.sin(math.radians(voltage_current_angle))) / 1000
    Reactive_Power = Reactive_Power
    return Reactive_Power


def Apparent_Power_calculate(voltage, current):
    Apparent_Power = (voltage * current / 1000)
    Apparent_Power = Apparent_Power
    return Apparent_Power


def Power_Factor_calculate(voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Power_Factor = math.cos(math.radians(voltage_current_angle))
    return Power_Factor


def System_Load_Nature_calculate(P_Sum, Q_Sum):
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


def Load_Nature_calculate(Vol_Angle, Cur_Angle):
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


def Power_standard_value_Calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang):
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
    Phase_A_P = Active_Power_calculate(Va, Ia, Va_ang, Ia_ang)
    Phase_A_Q = Reactive_Power_calculate(Va, Ia, Va_ang, Ia_ang)
    Phase_A_S = Apparent_Power_calculate(Va, Ia)
    Phase_A_PF = Power_Factor_calculate(Va_ang, Ia_ang)

    Phase_B_P = Active_Power_calculate(Vb, Ib, Vb_ang, Ib_ang)
    Phase_B_Q = Reactive_Power_calculate(Vb, Ib, Vb_ang, Ib_ang)
    Phase_B_S = Apparent_Power_calculate(Vb, Ib)
    Phase_B_PF = Power_Factor_calculate(Vb_ang, Ib_ang)

    Phase_C_P = Active_Power_calculate(Vc, Ic, Vc_ang, Ic_ang)
    Phase_C_Q = Reactive_Power_calculate(Vc, Ic, Vc_ang, Ic_ang)
    Phase_C_S = Apparent_Power_calculate(Vc, Ic)
    Phase_C_PF = Power_Factor_calculate(Vc_ang, Ic_ang)

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


# print(Power_standard_value_Calculate(69,69,69,3,3,3,0,240,120,330,210,90))


def Read_AcuRev4100_Power(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, User_Channel):
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
    Power = Power_standard_value_Calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang)
    Phase_A_P = Read_Phase_A_Active_Power(Power[0], times=40)
    Phase_A_Q = Read_Phase_A_Reactive_Power(Power[1], times=40)
    Phase_A_S = Read_Phase_A_Apparent_Power(Power[2], times=40)
    Phase_A_PF = Read_Phase_A_Power_Factor(Power[3], times=40)

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

    Phase_B_P = Read_Phase_B_Active_Power(Power[4], times=40)
    Phase_B_Q = Read_Phase_B_Reactive_Power(Power[5], times=40)
    Phase_B_S = Read_Phase_B_Apparent_Power(Power[6], times=40)
    Phase_B_PF = Read_Phase_B_Power_Factor(Power[7], times=40)

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

    Phase_C_P = Read_Phase_C_Active_Power(Power[8], times=40)
    Phase_C_Q = Read_Phase_C_Reactive_Power(Power[9], times=40)
    Phase_C_S = Read_Phase_C_Apparent_Power(Power[10], times=40)
    Phase_C_PF = Read_Phase_C_Power_Factor(Power[11], times=40)

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

    Sys_P = Read_System_Active_Power(Power[12], times=40)
    Sys_Q = Read_System_Reactive_Power(Power[13], times=40)
    Sys_S = Read_System_Apparent_Power(Power[14], times=40)
    Sys_PF = Read_System_Power_Factor(Power[15], times=40)

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

    Input1_P = Read_Input_Channel_1_Active_Power(Power[0], times=40)
    Input1_Q = Read_Input_Channel_1_Reactive_Power(Power[1], times=40)
    Input1_S = Read_Input_Channel_1_Apparent_Power(Power[2], times=40)
    Input1_PF = Read_Input_Channel_1_Power_Factor(Power[3], times=40)

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

    Input2_P = Read_Input_Channel_2_Active_Power(Power[4], times=40)
    Input2_Q = Read_Input_Channel_2_Reactive_Power(Power[5], times=40)
    Input2_S = Read_Input_Channel_2_Apparent_Power(Power[6], times=40)
    Input2_PF = Read_Input_Channel_2_Power_Factor(Power[7], times=40)

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

    Input3_P = Read_Input_Channel_3_Active_Power(Power[8], times=40)
    Input3_Q = Read_Input_Channel_3_Reactive_Power(Power[9], times=40)
    Input3_S = Read_Input_Channel_3_Apparent_Power(Power[10], times=40)
    Input3_PF = Read_Input_Channel_3_Power_Factor(Power[11], times=40)

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
        User1_P = Read_User_Channel_1_Active_Power(Power[12], times=40)
        User1_Q = Read_User_Channel_1_Reactive_Power(Power[13], times=40)
        User1_S = Read_User_Channel_1_Apparent_Power(Power[14], times=40)
        User1_PF = Read_User_Channel_1_Power_Factor(Power[15], times=40)

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
        User1_P = Read_User_Channel_1_Active_Power(Power[0], times=40)
        User1_Q = Read_User_Channel_1_Reactive_Power(Power[1], times=40)
        User1_S = Read_User_Channel_1_Apparent_Power(Power[2], times=40)
        User1_PF = Read_User_Channel_1_Power_Factor(Power[3], times=40)

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

        User2_P = Read_User_Channel_2_Active_Power(Power[4], times=40)
        User2_Q = Read_User_Channel_2_Reactive_Power(Power[5], times=40)
        User2_S = Read_User_Channel_2_Apparent_Power(Power[6], times=40)
        User2_PF = Read_User_Channel_2_Power_Factor(Power[7], times=40)

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

        User3_P = Read_User_Channel_3_Active_Power(Power[8], times=40)
        User3_Q = Read_User_Channel_3_Reactive_Power(Power[9], times=40)
        User3_S = Read_User_Channel_3_Apparent_Power(Power[10], times=40)
        User3_PF = Read_User_Channel_3_Power_Factor(Power[11], times=40)

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


# print(Read_AcuRev4100_Power(220, 220, 220, 4, 4, 4, 0, 240, 120, 180, 60, 300, 'user1,user2,user3'))


def Set_Reactive_Power_Calculation_Methodme(value):
    ret = ModbusClient.write_registers(address=4315, values=value, slave=1)
    if '(4315,1)' not in str(ret):
        logging.error('Set_Reactive_Power_Calculation_Methodme fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=4315, count=1, slave=1)
    if ret[0] == value:
        return True
    return False


def Set_Service_Configuration(value):
    ret = ModbusClient.write_registers(address=4162, values=value, slave=1)
    if '(4162,1)' not in str(ret):
        logging.error('Set_Service_Configuration fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=4162, count=1, slave=1)
    if ret[0] == value:
        return True
    return False


def Set_Clear_energy(value):
    '''
    :param value: 0: None 1: Clearing
    :return:
    '''
    ret = ModbusClient.write_registers(address=4400, values=value, slave=1)
    if '(4400,1)' not in str(ret):
        logging.error('Set_Clear_energy fail, ret is:{}'.format(ret))
        return False
    return True


# print(Set_Clear_energy(0))
def Set_Device_Reboot(value):
    '''
    :param value: 0: None 1: Clearing
    :return:
    '''
    ret = ModbusClient.write_registers(address=4420, values=value, slave=1)
    if '(4420,1)' not in str(ret):
        logging.error('Set_Device_Reboot fail, ret is:{}'.format(ret))
        return False
    return True


def Set_channle2_voltage_assignment(value):
    '''
    :param value: 1: Vb 2: Vc
    :return:
    '''
    ret = ModbusClient.write_registers(address=4175, values=value, slave=1)
    if '(4175,1)' not in str(ret):
        logging.error('Set_Channle2_voltage_assignment fail, ret is:{}'.format(ret))
        return False
    return True


def Set_channle3_voltage_assignment(value):
    '''
    :param value: 1: Vb 2: Vc
    :return:
    '''
    ret = ModbusClient.write_registers(address=4180, values=value, slave=1)
    if '(4180,1)' not in str(ret):
        logging.error('Set_channle3_voltage_assignment fail, ret is:{}'.format(ret))
        return False
    return True


def Set_Phase_Order(value):
    '''
    :param value: 0: ABC 1: ACB
    :return:
    '''
    ret = ModbusClient.write_registers(address=4316, values=value, slave=1)
    if '(4316,1)' not in str(ret):
        logging.error('Set_Phase_Order fail, ret is:{}'.format(ret))
        return False
    return True


def Read_Phase_A_Energy():
    Phase_A_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=12288, count=36, slave=1)
    if energy_charge == "resp is error":
        energy_charge: list = ModbusClient.read_measurement(address=12288, count=36, slave=1)
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


def Read_Phase_B_Energy():
    Phase_B_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=12324, count=36, slave=1)
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


def Read_Phase_C_Energy():
    Phase_C_Energy_list = []
    energy_charge: list = ModbusClient.read_measurement(address=12360, count=36, slave=1)
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


def Read_System_Energy():
    System_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=12396, count=36, slave=1)
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


def Read_Input_Channel_1_Energy():
    Input_Channel_1_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=12432, count=36, slave=1)
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


def Read_Input_Channel_2_Energy():
    Input_Channel_2_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=12468, count=36, slave=1)
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


def Read_Input_Channel_3_Energy():
    Input_Channel_3_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=12504, count=36, slave=1)
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


def Read_User_Channel_1_Energy():
    User_Channel_1_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=13296, count=36, slave=1)
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


def Read_User_Channel_2_Energy():
    User_Channel_2_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=13332, count=36, slave=1)
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


def Read_User_Channel_3_Energy():
    User_Channel_3_Energy = []
    energy_charge: list = ModbusClient.read_measurement(address=13368, count=36, slave=1)
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


def Energy_Standard_Value_Calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time):
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
    Power_standard_value_list = Power_standard_value_Calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang,
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


# print(Energy_Standard_Value_Calculate(100,100,100,15,15,15,0,240,120,0,120,240,10))

def Read_Energy_scale(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time, Service):
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
    Energy_Standard_Value_list = Energy_Standard_Value_Calculate(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang,
                                                                 Ib_ang, Ic_ang, Time)
    Phase_A_Energy_list = Read_Phase_A_Energy()
    Phase_B_Energy_list = Read_Phase_B_Energy()
    Phase_C_Energy_list = Read_Phase_C_Energy()
    System_Energy_list = Read_System_Energy()
    Input_Channel_1_Energy = Read_Input_Channel_1_Energy()
    Input_Channel_2_Energy = Read_Input_Channel_2_Energy()
    Input_Channel_3_Energy = Read_Input_Channel_3_Energy()
    User_Channel_1_Energy = Read_User_Channel_1_Energy()
    User_Channel_2_Energy = Read_User_Channel_2_Energy()
    User_Channel_3_Energy = Read_User_Channel_3_Energy()
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


def Read_Voltage_Positive_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9728, count=2, slave=1)
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


def Read_Voltage_Zero_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9730, count=2, slave=1)
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


def Read_Voltage_Negative_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9732, count=2, slave=1)
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


def Read_Voltage_Positive_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9734, count=2, slave=1)
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


def Read_Voltage_Zero_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9736, count=2, slave=1)
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


def Read_Voltage_Negative_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9738, count=2, slave=1)
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


def Read_Voltage_Unbalance_Factor_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=9740, count=2, slave=1)
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


def Read_User_Channel_1_Current_Positive_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11590, count=2, slave=1)
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


def Read_User_Channel_1_Current_Zero_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11592, count=2, slave=1)
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


def Read_User_Channel_1_Current_Negative_Sequence_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11594, count=2, slave=1)
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


def Read_User_Channel_1_Current_Positive_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11596, count=2, slave=1)
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


def Read_User_Channel_1_Current_Zero_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11598, count=2, slave=1)
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


def Read_User_Channel_1_Current_Negative_Sequence_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11600, count=2, slave=1)
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


def Read_User_Channel_1_Current_Unbalance_Factor_Magnitude(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=11602, count=2, slave=1)
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


def Read_Input_Channel_4_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8306, count=2, slave=1)
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


def Read_Input_Channel_5_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8320, count=2, slave=1)
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


def Read_Input_Channel_6_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8334, count=2, slave=1)
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


def Read_Input_Channel_7_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8348, count=2, slave=1)
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


def Read_Input_Channel_8_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8362, count=2, slave=1)
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


def Read_Input_Channel_9_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8376, count=2, slave=1)
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


def Read_Input_Channel_10_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8390, count=2, slave=1)
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


def Read_Input_Channel_11_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8404, count=2, slave=1)
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


def Read_Input_Channel_12_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8418, count=2, slave=1)
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


def Read_Input_Channel_13_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8432, count=2, slave=1)
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


def Read_Input_Channel_14_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8446, count=2, slave=1)
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


def Read_Input_Channel_15_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8460, count=2, slave=1)
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


def Read_Input_Channel_16_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8474, count=2, slave=1)
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


def Read_Input_Channel_17_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8488, count=2, slave=1)
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


def Read_Input_Channel_18_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8502, count=2, slave=1)
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


def Read_Input_Channel_19_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8516, count=2, slave=1)
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


def Read_Input_Channel_20_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8530, count=2, slave=1)
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


def Read_Input_Channel_21_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8544, count=2, slave=1)
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


def Read_Input_Channel_22_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8558, count=2, slave=1)
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


def Read_Input_Channel_23_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8572, count=2, slave=1)
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


def Read_Input_Channel_24_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8586, count=2, slave=1)
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


def Read_Phase_Sys_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8216, count=48, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
        value = [value[i] for i in range(len(value)) if
                 i not in {0, 1, 8, 9, 12, 13, 20, 21, 24, 25, 32, 33, 36, 37, 44, 45}]
        read_list = []
        # print(value)
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
    print(val_list)


# print(Read_Phase_Sys_Power([0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,0,0,0,0], 2))
def Read_Phase_Input_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8264, count=42, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
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


def Read_User_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
        time.sleep(0.2)
        value = ModbusClient.read_measurement(address=8600, count=36, slave=1)
        logging.info('Phase_A_Power ret is:{}'.format(value))
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


# print(Read_User_Power([0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0], 2))
def Power_standard_value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang,
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
        Phase_A_P = round(Active_Power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        Phase_A_Q = round(Reactive_Power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        Phase_A_S = round(Apparent_Power_calculate(Va, Ia), 10)
        Phase_A_PF = round(Power_Factor_calculate(Va_ang, Ia_ang), 10)

        Phase_B_P = round(Active_Power_calculate(Vb, Ib, Vb_ang, Ib_ang), 10)
        Phase_B_Q = round(Reactive_Power_calculate(Vb, Ib, Vb_ang, Ib_ang), 10)
        Phase_B_S = round(Apparent_Power_calculate(Vb, Ib), 10)
        Phase_B_PF = round(Power_Factor_calculate(Vb_ang, Ib_ang), 10)

        Phase_C_P = round(Active_Power_calculate(Vc, Ic, Vc_ang, Ic_ang), 10)
        Phase_C_Q = round(Reactive_Power_calculate(Vc, Ic, Vc_ang, Ic_ang), 10)
        Phase_C_S = round(Apparent_Power_calculate(Vc, Ic), 10)
        Phase_C_PF = round(Power_Factor_calculate(Vc_ang, Ic_ang), 10)

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
        input_1_P = round(Active_Power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        input_1_Q = round(Reactive_Power_calculate(Va, Ia, Va_ang, Ia_ang), 10)
        input_1_S = round(Apparent_Power_calculate(Va, Ia), 10)
        input_1_PF = round(Power_Factor_calculate(Va_ang, Ia_ang), 10)

        input_2_P = round(Active_Power_calculate(Va, Ib, Va_ang, Ib_ang), 10)
        input_2_Q = round(Reactive_Power_calculate(Va, Ib, Va_ang, Ib_ang), 10)
        input_2_S = round(Apparent_Power_calculate(Va, Ib), 10)
        input_2_PF = round(Power_Factor_calculate(Va_ang, Ib_ang), 10)

        input_3_P = round(Active_Power_calculate(Va, Ic, Va_ang, Ic_ang), 10)
        input_3_Q = round(Reactive_Power_calculate(Va, Ic, Va_ang, Ic_ang), 10)
        input_3_S = round(Apparent_Power_calculate(Va, Ic), 10)
        input_3_PF = round(Power_Factor_calculate(Va_ang, Ic_ang), 10)

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


# print(Power_standard_value_Calculate_new(69,69,69,3,3,3,0,0,0,0,0,0,"1E1p2w"))
def Read_AcuRev4100_Power_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, User_Channel):
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
    Power = Power_standard_value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang,
                                               User_Channel)
    # print(Power[16:])
    Read_Phase_Power = Read_Phase_Sys_Power(Power, times=40)
    Read_Input_Power = []
    if User_Channel == "3E3p4w":
        Read_Input_Power = Read_Phase_Input_Power(Power[:12], times=40)
    if User_Channel == "1E1p2w":
        Read_Input_Power = Read_Phase_Input_Power(Power[16:], times=40)
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
        User1_P = Read_User_Channel_1_Active_Power(Power[12], times=40)
        User1_Q = Read_User_Channel_1_Reactive_Power(Power[13], times=40)
        User1_S = Read_User_Channel_1_Apparent_Power(Power[14], times=40)
        User1_PF = Read_User_Channel_1_Power_Factor(Power[15], times=40)

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
    # if User_Channel == '1E1p2w':
    #     Read_User_Power_list = Read_User_Power(Power[16:], times=40)
    #     User1_P = round(Read_User_Power_list[0], 6)
    #     User1_Q = Read_User_Power_list[1]
    #     User1_S = Read_User_Power_list[2]
    #     User1_PF = Read_User_Power_list[3]
    #
    #     if Power[16] != 0:
    #         scale_User1_P = abs((User1_P - Power[16]) / Power[16])
    #     else:
    #         if Power[16] == User1_P == 0 or User1_P < 0.001:
    #             scale_User1_P = 0
    #         else:
    #             scale_User1_P = 'null'
    #
    #     if Power[17] != 0:
    #         scale_User1_Q = abs((User1_Q - Power[17]) / Power[17])
    #     else:
    #         if Power[17] == User1_Q == 0 or User1_Q < 0.001:
    #             scale_User1_Q = 0
    #         else:
    #             scale_User1_Q = 'null'
    #
    #     if Power[18] != 0:
    #         scale_User1_S = abs((User1_S - Power[18]) / Power[18])
    #     else:
    #         if Power[18] == User1_S == 0 or User1_S < 0.001:
    #             scale_User1_S = 0
    #         else:
    #             scale_User1_S = 'null'
    #
    #     if Power[19] != 0:
    #         scale_User1_PF = abs((User1_PF - Power[19]) / Power[19])
    #     else:
    #         if Power[19] == User1_PF == 0 or User1_PF < 0.001:
    #             scale_User1_PF = 0
    #         else:
    #             scale_User1_PF = 'null'
    #
    #     User2_P = Read_User_Power_list[4]
    #     User2_Q = Read_User_Power_list[5]
    #     User2_S = Read_User_Power_list[6]
    #     User2_PF = Read_User_Power_list[7]
    #
    #     if Power[20] != 0:
    #         scale_User2_P = abs((User2_P - Power[20]) / Power[20])
    #     else:
    #         if Power[20] == User2_P == 0 or User2_P < 0.001:
    #             scale_User2_P = 0
    #         else:
    #             scale_User2_P = 'null'
    #
    #     if Power[21] != 0:
    #         scale_User2_Q = abs((User2_Q - Power[21]) / Power[21])
    #     else:
    #         if Power[21] == User2_Q == 0 or User2_Q < 0.001:
    #             scale_User2_Q = 0
    #         else:
    #             scale_User2_Q = 'null'
    #
    #     if Power[22] != 0:
    #         scale_User2_S = abs((User2_S - Power[22]) / Power[22])
    #     else:
    #         if Power[22] == User2_S == 0 or User2_S < 0.001:
    #             scale_User2_S = 0
    #         else:
    #             scale_User2_S = 'null'
    #
    #     if Power[23] != 0:
    #         scale_User2_PF = abs((User2_PF - Power[23]) / Power[23])
    #     else:
    #         if Power[23] == User2_PF == 0 or User2_PF < 0.001:
    #             scale_User2_PF = 0
    #         else:
    #             scale_User2_PF = 'null'
    #
    #     User3_P = Read_User_Power_list[8]
    #     User3_Q = Read_User_Power_list[9]
    #     User3_S = Read_User_Power_list[10]
    #     User3_PF = Read_User_Power_list[11]
    #
    #     if Power[24] != 0:
    #         scale_User3_P = abs((User3_P - Power[24]) / Power[24])
    #     else:
    #         if Power[24] == User3_P == 0 or User3_P < 0.001:
    #             scale_User3_P = 0
    #         else:
    #             scale_User3_P = 'null'
    #     if Power[25] != 0:
    #         scale_User3_Q = abs((User3_Q - Power[25]) / Power[25])
    #     else:
    #         if Power[25] == User3_Q == 0 or User3_Q < 0.001:
    #             scale_User3_Q = 0
    #         else:
    #             scale_User3_Q = 'null'
    #     if Power[26] != 0:
    #         scale_User3_S = abs((User3_S - Power[26]) / Power[26])
    #     else:
    #         if Power[26] == User3_S == 0 or User3_S < 0.001:
    #             scale_User3_S = 0
    #         else:
    #             scale_User3_S = 'null'
    #     if Power[27] != 0:
    #         scale_User3_PF = abs((User3_PF - Power[27]) / Power[27])
    #     else:
    #         if Power[27] == User3_PF == 0 or User3_PF < 0.001:
    #             scale_User3_PF = 0
    #         else:
    #             scale_User3_PF = 'null'
    #     power_list.extend(
    #         [[User1_P, scale_User1_P], [User1_Q, scale_User1_Q], [User1_S, scale_User1_S], [User1_PF, scale_User1_PF],
    #          [User2_P, scale_User2_P], [User2_Q, scale_User2_Q], [User2_S, scale_User2_S], [User2_PF, scale_User2_PF],
    #          [User3_P, scale_User3_P], [User3_Q, scale_User3_Q], [User3_S, scale_User3_S], [User3_PF, scale_User3_PF]])
    return power_list
# print(Read_AcuRev4100_Power_new(69,69,69,3,3,3,0,0,0,0,0,0,"1E1p2w"))

def Energy_Standard_Value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time,
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
        Power_standard_value_list = Power_standard_value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
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
        Power_standard_value_list = Power_standard_value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
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


# print(Energy_Standard_Value_Calculate_new(120,120,120,3,15,20,0,240,120,330,330,330,20,'1E1p2w'))

def Read_Energy_scale_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang, Ia_ang, Ib_ang, Ic_ang, Time, Service):
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
    Energy_Standard_Value_list = Energy_Standard_Value_Calculate_new(Va, Vb, Vc, Ia, Ib, Ic, Va_ang, Vb_ang, Vc_ang,
                                                                     Ia_ang,
                                                                     Ib_ang, Ic_ang, Time, Service)
    Phase_A_Energy_list = Read_Phase_A_Energy()
    Phase_B_Energy_list = Read_Phase_B_Energy()
    Phase_C_Energy_list = Read_Phase_C_Energy()
    System_Energy_list = Read_System_Energy()
    Input_Channel_1_Energy = Read_Input_Channel_1_Energy()
    Input_Channel_2_Energy = Read_Input_Channel_2_Energy()
    Input_Channel_3_Energy = Read_Input_Channel_3_Energy()
    User_Channel_1_Energy = Read_User_Channel_1_Energy()
    User_Channel_2_Energy = Read_User_Channel_2_Energy()
    User_Channel_3_Energy = Read_User_Channel_3_Energy()
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
            Energy_scale = 'null'
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
                Energy_scale = 'null'
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
        for i in range(36, 63):
            if Energy_Standard_Value_list[i] != 0:
                Energy_scale = abs(
                    (Read_Energy_list[i] - Energy_Standard_Value_list[i]) / Energy_Standard_Value_list[i])
                Energy_scale_list.append(Energy_scale)
            elif Energy_Standard_Value_list[i] == Read_Energy_list[i] == 0 or Read_Energy_list[i] < 0.001:
                Energy_scale = 0
                Energy_scale_list.append(Energy_scale)
            else:
                Energy_scale = 'null'
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
                Energy_scale = 'null'
                Energy_scale_list.append(Energy_scale)
    return Read_Energy_list, Energy_scale_list


# print(Read_AcuRev4100_Power_new(100, 100, 100, 5, 5, 5, 0, 240, 120, 0, 240, 120, 'user1'))
def wait_minutes(wait_times):
    """等待指定分钟后输出"""
    time.sleep(wait_times * 60)  # 转换成秒
    print(f"等待 {wait_times} 分钟结束！")


def hold_rs485_connect(hold_time):
    """每 5 分钟读取一次 Modbus 数据"""
    t = int((hold_time / 3) - 1)  # 计算读取次数
    print(t)
    for i in range(t):
        time.sleep(180)  # 3 分钟
        value = ModbusClient.read_measurement(address=8192, count=2, slave=1)
        print(f"第 {i + 1} 次读取数据: {value},RS485 连接正常")


# thread_a = threading.Thread(target=wait_minutes, args=(20,))
# thread_b = threading.Thread(target=hold_rs485_connect, args=(20,))
# t = time.time()
# print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t)) + f".{int((t * 1000) % 1000):03d}")
# # 启动线程
# thread_a.start()
# thread_b.start()
#
# # 等待线程完成
# thread_a.join()
# thread_b.join()
#
# print("主线程: 所有任务完成，退出程序！")
# t = time.time()
# print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t)) + f".{int((t * 1000) % 1000):03d}")


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

# # 示例：基波电流为 1 A，第2次谐波为10%
# base_current = 1
# harmonics = {
#     2: 10  # 表示第2次谐波占基波的10%
# }
