from comm.modbus_rtu_tcp import *
from comm.source_control import *
from tools.log import Log
import xlwt
from tools.excel_operate import data_read
from modbus_config import modbus_config
import math

# volt_cur_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])

ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def read_frequency(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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


def Read_Phase_C_Voltage(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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
        return val_list[-1]
    return val_list[0]


def Read_Phase_B_Voltage_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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
        return val_list[-1]
    return val_list[0]


def Read_Phase_C_Voltage_Angle(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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
        return val_list[-1]
    return val_list[0]


def Read_Phase_A_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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


def Read_Phase_A_Active_Power(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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
        return val_list[-1]
    return val_list[0]


def Read_Input_Channel_2_Current(standard_value, times=1):
    # ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    val_list = []
    for v in range(times):
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
    Active_Power = round(Active_Power, 3)
    return Active_Power


def Reactive_Power_calculate(voltage, current, voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Reactive_Power = (voltage * current * math.sin(math.radians(voltage_current_angle))) / 1000
    Reactive_Power = round(Reactive_Power, 3)
    return Reactive_Power


def Apparent_Power_calculate(voltage, current):
    Apparent_Power = (voltage * current / 1000)
    Apparent_Power = round(Apparent_Power, 3)
    return Apparent_Power


def Power_Factor_calculate(voltage_angle, current_angle):
    voltage_current_angle = voltage_angle - current_angle
    Power_Factor = round(math.cos(math.radians(voltage_current_angle)), 3)
    return Power_Factor


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
    Phase_A_P = Read_Phase_A_Active_Power(Power[0], 10)
    Phase_A_Q = Read_Phase_A_Reactive_Power(Power[1], 10)
    Phase_A_S = Read_Phase_A_Apparent_Power(Power[2], 10)
    Phase_A_PF = Read_Phase_A_Power_Factor(Power[3], 10)

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

    Phase_B_P = Read_Phase_B_Active_Power(Power[4], 10)
    Phase_B_Q = Read_Phase_B_Reactive_Power(Power[5], 10)
    Phase_B_S = Read_Phase_B_Apparent_Power(Power[6], 10)
    Phase_B_PF = Read_Phase_B_Power_Factor(Power[7], 10)

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

    Phase_C_P = Read_Phase_C_Active_Power(Power[8], 10)
    Phase_C_Q = Read_Phase_C_Reactive_Power(Power[9], 10)
    Phase_C_S = Read_Phase_C_Apparent_Power(Power[10], 10)
    Phase_C_PF = Read_Phase_C_Power_Factor(Power[11], 10)

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

    Sys_P = Read_System_Active_Power(Power[12], 10)
    Sys_Q = Read_System_Reactive_Power(Power[13], 10)
    Sys_S = Read_System_Apparent_Power(Power[14], 10)
    Sys_PF = Read_System_Power_Factor(Power[15], 10)

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

    Input1_P = Read_Input_Channel_1_Active_Power(Power[0], 10)
    Input1_Q = Read_Input_Channel_1_Reactive_Power(Power[1], 10)
    Input1_S = Read_Input_Channel_1_Apparent_Power(Power[2], 10)
    Input1_PF = Read_Input_Channel_1_Power_Factor(Power[3], 10)

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

    Input2_P = Read_Input_Channel_2_Active_Power(Power[4], 10)
    Input2_Q = Read_Input_Channel_2_Reactive_Power(Power[5], 10)
    Input2_S = Read_Input_Channel_2_Apparent_Power(Power[6], 10)
    Input2_PF = Read_Input_Channel_2_Power_Factor(Power[7], 10)

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

    Input3_P = Read_Input_Channel_3_Active_Power(Power[8], 10)
    Input3_Q = Read_Input_Channel_3_Reactive_Power(Power[9], 10)
    Input3_S = Read_Input_Channel_3_Apparent_Power(Power[10], 10)
    Input3_PF = Read_Input_Channel_3_Power_Factor(Power[11], 10)

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
        User1_P = Read_User_Channel_1_Active_Power(Power[12], 10)
        User1_Q = Read_User_Channel_1_Reactive_Power(Power[13], 10)
        User1_S = Read_User_Channel_1_Apparent_Power(Power[14], 10)
        User1_PF = Read_User_Channel_1_Power_Factor(Power[15], 10)

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
        User1_P = Read_User_Channel_1_Active_Power(Power[0], 10)
        User1_Q = Read_User_Channel_1_Reactive_Power(Power[1], 10)
        User1_S = Read_User_Channel_1_Apparent_Power(Power[2], 10)
        User1_PF = Read_User_Channel_1_Power_Factor(Power[3], 10)

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

        User2_P = Read_User_Channel_2_Active_Power(Power[4], 10)
        User2_Q = Read_User_Channel_2_Reactive_Power(Power[5], 10)
        User2_S = Read_User_Channel_2_Apparent_Power(Power[6], 10)
        User2_PF = Read_User_Channel_2_Power_Factor(Power[7], 10)

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

        User3_P = Read_User_Channel_3_Active_Power(Power[8], 10)
        User3_Q = Read_User_Channel_3_Reactive_Power(Power[9], 10)
        User3_S = Read_User_Channel_3_Apparent_Power(Power[10], 10)
        User3_PF = Read_User_Channel_3_Power_Factor(Power[11], 10)

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
        print(ret)
        return False
    ret = ModbusClient.read_measurement(address=4315, count=1, slave=1)
    print(ret)
    if ret[0] == value:
        return True
    return False
