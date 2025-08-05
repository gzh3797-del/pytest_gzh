import logging

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from tools.excel_operate import dcpara_4100addr_get
import struct

dc_para_addr = dcpara_4100addr_get(r'/comm/test_data/AcuDC300.xlsx', 'Readings')
basic_configuration = dcpara_4100addr_get(r'/comm/test_data/AcuDC300.xlsx', 'Basic Configuration')
System_Infomation = dcpara_4100addr_get(r'/comm/test_data/AcuDC300.xlsx', 'System Infomation')
Calibration = dcpara_4100addr_get(r'/comm/test_data/AcuDC300.xlsx', 'Calibration')


def dec_to_signed_decimal(dec_str, bit_width):
    if dec_str & (1 << (bit_width - 1)):
        # 计算补码形式的负数
        dec_str -= (1 << bit_width)

    return dec_str


def read_voltage_measurement():
    client = ModbusRtuOrTcp()
    voltage: list = client.read_measurement(address=dc_para_addr['V(Measured) float32']['Start(Dec)'],
                                            count=dc_para_addr['V(Measured) float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]).replace('0x', '').zfill(4) + hex(voltage[1]).replace('0x', '').zfill(4)
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_voltage_measurement_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['V(Measured) int16']['Start(Dec)'],
                                        count=dc_para_addr['V(Measured) int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_voltage_measu_or_comp_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['V(Measured or Compensated) int16']['Start(Dec)'],
                                        count=dc_para_addr['V(Measured or Compensated) int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_current_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Current int16']['Start(Dec)'],
                                        count=dc_para_addr['Current int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_power_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Power int16']['Start(Dec)'],
                                        count=dc_para_addr['Power int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_voltage_ripple_factor_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Voltage Ripple Factor int16']['Start(Dec)'],
                                        count=dc_para_addr['Voltage Ripple Factor int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16) / 100
    return ret


def read_current_ripple_factor_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Current Ripple Factor int16']['Start(Dec)'],
                                        count=dc_para_addr['Current Ripple Factor int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16) / 100
    return ret


def read_demand_current_import_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Demand Current import int16']['Start(Dec)'],
                                        count=dc_para_addr['Demand Current import int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_demand_current_export_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Demand Current export int16']['Start(Dec)'],
                                        count=dc_para_addr['Demand Current export int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_demand_power_import_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Demand Power import int16']['Start(Dec)'],
                                        count=dc_para_addr['Demand Power import int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_demand_power_export_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['Demand Power export int16']['Start(Dec)'],
                                        count=dc_para_addr['Demand Power export int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_voltage_comp_client(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=dc_para_addr['V(Compensated) int16']['Start(Dec)'],
                                        count=dc_para_addr['V(Compensated) int16']['Reg'], slave=1)
    client.close()
    ret = dec_to_signed_decimal(ret[0], 16)
    return ret


def read_voltage_measu_or_comp(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['V(Measured or Compensated) float32']['Start(Dec)'],
                                            count=dc_para_addr['V(Measured or Compensated) float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_voltage_compensated(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['V(Compensated) float32']['Start(Dec)'],
                                            count=dc_para_addr['V(Compensated) float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_current_measurement():
    client = ModbusRtuOrTcp()
    voltage: list = client.read_measurement(address=dc_para_addr['Current float32']['Start(Dec)'],
                                            count=dc_para_addr['Current float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]).replace('0x', '').zfill(4) + hex(voltage[1]).replace('0x', '').zfill(4)
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_power_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Power float32']['Start(Dec)'],
                                            count=dc_para_addr['Power float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_voltage_ripple_factor_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Voltage Ripple Factor float32']['Start(Dec)'],
                                            count=dc_para_addr['Voltage Ripple Factor float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_current_ripple_factor_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Current Ripple Factor float32']['Start(Dec)'],
                                            count=dc_para_addr['Current Ripple Factor float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_demand_current_import_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Demand Current import float32']['Start(Dec)'],
                                            count=dc_para_addr['Demand Current import float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_demand_current_export_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Demand Current export float32']['Start(Dec)'],
                                            count=dc_para_addr['Demand Current export float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_demand_power_import_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Demand Power import float32']['Start(Dec)'],
                                            count=dc_para_addr['Demand Power import float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_demand_power_export_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Demand Power export float32']['Start(Dec)'],
                                            count=dc_para_addr['Demand Power export float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_import_energy_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Import Energy double']['Start(Dec)'],
                                                  count=dc_para_addr['Import Energy double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['import_energy'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_energy_charge(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    energy_charge: list = client.read_measurement(address=dc_para_addr['Import Energy double']['Start(Dec)'],
                                                  count=dc_para_addr['Import Energy double']['Reg'] * 8, slave=1)
    client.close()
    import_energy = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace('0x', '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(import_energy, 16)
    import_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    export_energy = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace('0x', '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(export_energy, 16)
    export_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    net_energy = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace('0x', '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(net_energy, 16)
    net_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    total_energy = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(energy_charge[13]).replace('0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(total_energy, 16)
    total_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    import_charge = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(energy_charge[17]).replace('0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(import_charge, 16)
    import_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    export_charge = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(energy_charge[21]).replace('0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(export_charge, 16)
    export_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    net_charge = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(energy_charge[25]).replace('0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(net_charge, 16)
    net_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    total_charge = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(energy_charge[29]).replace('0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(total_charge, 16)
    total_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return (import_energy, export_energy, net_energy, total_energy, import_charge, export_charge, net_charge,
            total_charge)


def read_export_energy_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Export Energy double']['Start(Dec)'],
                                                  count=dc_para_addr['Export Energy double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['export_energy'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_net_energy_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Net Energy double']['Start(Dec)'],
                                                  count=dc_para_addr['Net Energy double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['net_energy'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_total_energy_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Total Energy double']['Start(Dec)'],
                                                  count=dc_para_addr['Total Energy double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['total_energy'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_import_charge_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Import Charge double']['Start(Dec)'],
                                                  count=dc_para_addr['Import Charge double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['import_charge'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_export_charge_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Export Charge double']['Start(Dec)'],
                                                  count=dc_para_addr['Export Charge double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['export_charge'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_net_charge_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Net Charge double']['Start(Dec)'],
                                                  count=dc_para_addr['Net Charge double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['net_charge'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def read_total_charge_measurement(conn_mode, ret=None):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    import_energy: list = client.read_measurement(address=dc_para_addr['Total Charge double']['Start(Dec)'],
                                                  count=dc_para_addr['Total Charge double']['Reg'], slave=1)
    client.close()
    reg = hex(import_energy[0]) + hex(import_energy[1]) + hex(import_energy[1]) + hex(import_energy[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    ret['total_charge'] = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return float(struct.unpack('<d', struct.pack('<Q', integer_num))[0])


def get_real_time_clock(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    year = client.read_measurement(address=basic_configuration['Year word']['Start(Dec)'],
                                   count=int(basic_configuration['Year word']['Reg']), slave=1)
    month = client.read_measurement(address=basic_configuration['Month word']['Start(Dec)'],
                                    count=int(basic_configuration['Month word']['Reg']), slave=1)
    day = client.read_measurement(address=basic_configuration['Day word']['Start(Dec)'],
                                  count=int(basic_configuration['Day word']['Reg']), slave=1)
    hour = client.read_measurement(address=basic_configuration['Hour word']['Start(Dec)'],
                                   count=int(basic_configuration['Hour word']['Reg']), slave=1)
    minute = client.read_measurement(address=basic_configuration['Minute word']['Start(Dec)'],
                                     count=int(basic_configuration['Minute word']['Reg']), slave=1)
    second = client.read_measurement(address=basic_configuration['Second word']['Start(Dec)'],
                                     count=int(basic_configuration['Second word']['Reg']), slave=1)

    real_time_clock = str(year[0]) + '-' + str(month[0]) + '-' + str(day[0]) + ' ' + str(hour[0]) + ':' + str(
        minute[0]) + ':' + str(second[0])
    client.close()
    return real_time_clock


def get_meter_password(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Meter password word']['Start(Dec)'],
                                  count=int(basic_configuration['Meter password word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_rs485_baudrate(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['RS485 baud rate word']['Start(Dec)'],
                                  count=int(basic_configuration['RS485 baud rate word']['Reg']), slave=1)
    client.close()
    if type(ret) is not list:
        logging.error('get rs485 baudrate fail, ret is:{}'.format(ret))
    return ret[0]


def get_rs485_parity(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['RS485 parity word']['Start(Dec)'],
                                  count=int(basic_configuration['RS485 parity word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_dhcp_enable(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['DHCP enable word']['Start(Dec)'],
                                  count=int(basic_configuration['DHCP enable word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_ip_address(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['IP address 1st byte (high)\nIP address 2nd byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['IP address 1st byte (high)\nIP address 2nd byte (low) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['IP address 3rd byte (high)\nIP address 4th byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['IP address 3rd byte (high)\nIP address 4th byte (low) word']['Reg']),
        slave=1)
    ret: str = hex(ret[0]).replace('0x', '').zfill(4)
    ret1: str = hex(ret1[0]).replace('0x', '').zfill(4)
    ret: str = str(int(ret[0:2], 16)) + '.' + str(int(ret[2:4], 16)) + '.' + str(int(ret1[0:2], 16)) + '.' + str(
        int(ret1[2:4], 16))
    client.close()
    return ret


def get_subnet_mask(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['Subnet mask 1st byte (high)\nSubnet mask 2nd byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['Subnet mask 1st byte (high)\nSubnet mask 2nd byte (low) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['Subnet mask 3rd byte (high)\nSubnet mask 4th byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['Subnet mask 3rd byte (high)\nSubnet mask 4th byte (low) word']['Reg']),
        slave=1)
    ret: str = hex(ret[0]).replace('0x', '').zfill(4)
    ret1: str = hex(ret1[0]).replace('0x', '').zfill(4)
    ret: str = str(int(ret[0:2], 16)) + '.' + str(int(ret[2:4], 16)) + '.' + str(int(ret1[0:2], 16)) + '.' + str(
        int(ret1[2:4], 16))
    client.close()
    return ret


def get_gateway(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['Gateway 1st byte (high)\nGateway 2nd byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['Gateway 1st byte (high)\nGateway 2nd byte (low) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['Gateway 3rd byte (high)\nGateway 4th byte (low) word']['Start(Dec)'],
        count=int(basic_configuration['Gateway 3rd byte (high)\nGateway 4th byte (low) word']['Reg']),
        slave=1)

    ret: str = hex(ret[0]).replace('0x', '').zfill(4)
    ret1: str = hex(ret1[0]).replace('0x', '').zfill(4)
    ret: str = str(int(ret[0:2], 16)) + '.' + str(int(ret[2:4], 16)) + '.' + str(int(ret1[0:2], 16)) + '.' + str(
        int(ret1[2:4], 16))
    client.close()
    return ret


def get_dns_primary_server(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['DNS primary server 1st byte (high)\nDNS primary server 2nd byte (low) word'][
            'Start(Dec)'],
        count=int(
            basic_configuration['DNS primary server 1st byte (high)\nDNS primary server 2nd byte (low) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['DNS primary server 3rd byte (high)\nDNS primary server 4th byte (low) word'][
            'Start(Dec)'],
        count=int(
            basic_configuration['DNS primary server 3rd byte (high)\nDNS primary server 4th byte (low) word']['Reg']),
        slave=1)
    ret: str = hex(ret[0]).replace('0x', '').zfill(4)
    ret1: str = hex(ret1[0]).replace('0x', '').zfill(4)
    ret: str = str(int(ret[0:2], 16)) + '.' + str(int(ret[2:4], 16)) + '.' + str(int(ret1[0:2], 16)) + '.' + str(
        int(ret1[2:4], 16))
    client.close()
    return ret


def get_dns_secondary_server(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['DNS secondary server 1st byte (high)\nDNS secondary server 2nd byte (low) word'][
            'Start(Dec)'],
        count=int(
            basic_configuration['DNS secondary server 1st byte (high)\nDNS secondary server 2nd byte (low) word'][
                'Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['DNS secondary server 3rd byte (high)\nDNS secondary server 4th byte (low) word'][
            'Start(Dec)'],
        count=int(
            basic_configuration['DNS secondary server 3rd byte (high)\nDNS secondary server 4th byte (low) word'][
                'Reg']),
        slave=1)
    ret: str = hex(ret[0]).replace('0x', '').zfill(4)
    ret1: str = hex(ret1[0]).replace('0x', '').zfill(4)
    ret: str = str(int(ret[0:2], 16)) + '.' + str(int(ret[2:4], 16)) + '.' + str(int(ret1[0:2], 16)) + '.' + str(
        int(ret1[2:4], 16))
    client.close()
    return ret


def get_slave_id(conn_mode, slave=1):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Modbus Slave ID word']['Start(Dec)'],
                                  count=int(basic_configuration['Modbus Slave ID word']['Reg']), slave=slave)
    client.close()
    return ret[0]


def get_rtu_enable(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Modbus RTU Enable word']['Start(Dec)'],
                                  count=int(basic_configuration['Modbus RTU Enable word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_tcp_enable(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Modbus TCP Enable word']['Start(Dec)'],
                                  count=int(basic_configuration['Modbus TCP Enable word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_tcp_port(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Modbus TCP port word']['Start(Dec)'],
                                  count=int(basic_configuration['Modbus TCP port word']['Reg']), slave=1)
    client.close()
    logging.info('get_tcp_port ret is:{}'.format(ret))
    return ret[0]


def get_pt1(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['PT1 word']['Start(Dec)'],
                                  count=4, slave=1)
    client.close()
    return ret


def get_pt2(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['PT2 word']['Start(Dec)'],
                                  count=int(basic_configuration['PT2 word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_ct1(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['CT1 word']['Start(Dec)'],
                                  count=int(basic_configuration['CT1 word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_ct2(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['CT2 word']['Start(Dec)'],
                                  count=4, slave=1)
    client.close()
    return ret


def get_demand_calculation_method(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Demand calculation method word']['Start(Dec)'],
                                  count=int(basic_configuration['Demand calculation method word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_demand_window_time(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Demand window time word']['Start(Dec)'],
                                  count=int(basic_configuration['Demand window time word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_demand_update_period(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Demand update period word']['Start(Dec)'],
                                  count=int(basic_configuration['Demand update period word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_energy_pulse_para(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Energy pulse parameter word']['Start(Dec)'],
                                  count=int(basic_configuration['Energy pulse parameter word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_energy_pulse_constant(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=basic_configuration['Energy pulse constant word']['Start(Dec)'],
                                        count=basic_configuration['Energy pulse constant word']['Reg'], slave=1)
    client.close()
    ret = round((ret[0] * 65536 + ret[1]) / 1000, 3)
    return ret


def get_backlight_time(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Backlight time word']['Start(Dec)'],
                                  count=int(basic_configuration['Backlight time word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_seal_status(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Seal status word']['Start(Dec)'],
                                  count=int(basic_configuration['Seal status word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_device_run_time(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['Device run time (high 16 bits) word']['Start(Dec)'],
        count=int(basic_configuration['Device run time (high 16 bits) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['Device run time (low 16 bits) Word']['Start(Dec)'],
        count=int(basic_configuration['Device run time (low 16 bits) Word']['Reg']), slave=1)
    client.close()
    ret: float = round(int(str(ret[0]) + str(ret1[0])) / 3600, 2)
    return ret


def get_device_load_time(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(
        address=basic_configuration['Device load time (high 16 bits) word']['Start(Dec)'],
        count=int(basic_configuration['Device load time (high 16 bits) word']['Reg']),
        slave=1)
    ret1: list = client.read_measurement(
        address=basic_configuration['Device load time (low 16 bits) word']['Start(Dec)'],
        count=int(basic_configuration['Device load time (low 16 bits) word']['Reg']),
        slave=1)
    client.close()
    ret: float = round(int(str(ret[0]) + str(ret1[0])) / 3600, 2)
    return ret


def get_enable_cable_loss_compensation(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Enable cable loss compensation word']['Start(Dec)'],
                                  count=int(basic_configuration['Enable cable loss compensation word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_cable_resistance(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=basic_configuration['Cable  resistance word']['Start(Dec)'],
                                  count=int(basic_configuration['Cable  resistance word']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_firmware_version(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Firmware version']['Start(Dec)'],
                                        count=int(System_Infomation['Firmware version']['Reg']), slave=1)
    client.close()
    hex_ret = str((hex(ret[0])[2:6])) + str((hex(ret[1])[2:6]))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret


def get_firmware_release_date(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Firmware release date']['Start(Dec)'],
                                        count=int(System_Infomation['Firmware release date']['Reg']), slave=1)
    client.close()
    hex_ret = str((hex(ret[0])[2:6])) + str((hex(ret[1])[2:6])) + str((hex(ret[2])[2:6]))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    ret: str = '20' + ret[0:2] + '-' + ret[2:4] + '-' + ret[4:6]
    return ret


def get_firmware_patch_number(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Firmware patch number']['Start(Dec)'],
                                        count=int(System_Infomation['Firmware patch number']['Reg']), slave=1)
    client.close()
    hex_ret = str((hex(ret[0])[2:6]))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret


def get_serial_number(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Serial Number']['Start(Dec)'],
                                        count=int(System_Infomation['Serial Number']['Reg']), slave=1)
    client.close()
    hex_ret = (str((hex(ret[0])[2:6])) + str((hex(ret[1])[2:6])) + str((hex(ret[2])[2:6])) + str((hex(ret[3])[2:6]))
               + str((hex(ret[4])[2:6])) + str((hex(ret[5])[2:6])) + str((hex(ret[6])[2:6])))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret


def get_hardware_version(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Hardware version']['Start(Dec)'],
                                        count=int(System_Infomation['Hardware version']['Reg']), slave=1)
    client.close()
    hex_ret = str((hex(ret[0])[2:6])) + str((hex(ret[1])[2:6]))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret


def get_function_model_type(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Function Model Type']['Start(Dec)'],
                                        count=int(System_Infomation['Function Model Type']['Reg']), slave=1)
    client.close()
    # hex_ret = str((hex(ret[0])[2:6]))
    # ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret[0]


def get_voltage_input_type(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Voltage Input Type']['Start(Dec)'],
                                        count=int(System_Infomation['Voltage Input Type']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_current_input_type(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Current Input Type']['Start(Dec)'],
                                        count=int(System_Infomation['Current Input Type']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_power_supply_type(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Power Supply Type']['Start(Dec)'],
                                        count=int(System_Infomation['Power Supply Type']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_mac_address(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['MAC address']['Start(Dec)'],
                                        count=int(System_Infomation['MAC address']['Reg']), slave=1)
    client.close()
    hex_ret = (str((hex(ret[0])[2:6])) + str((hex(ret[1])[2:6])) + str((hex(ret[2])[2:6])) + str((hex(ret[3])[2:6]))
               + str((hex(ret[4])[2:6])) + str((hex(ret[5])[2:6])) + str((hex(ret[6])[2:6])) + str((hex(ret[7])[2:6]))
               + str((hex(ret[8])[2:6])) + str((hex(ret[9])[2:6])) + str((hex(ret[10])[2:6])) + str((hex(ret[11])[2:6]))
               + str((hex(ret[12])[2:6])) + str((hex(ret[13])[2:6])) + str((hex(ret[14])[2:6])) + str(
                (hex(ret[15])[2:6])))
    ret: str = bytes.fromhex(hex_ret).decode('ascii')
    return ret


def get_device_type(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Device Type']['Start(Dec)'],
                                        count=int(System_Infomation['Device Type']['Reg']), slave=1)
    client.close()
    return ret[0]


def get_reserved(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret: list = client.read_measurement(address=System_Infomation['Reserved']['Start(Dec)'],
                                        count=int(System_Infomation['Reserved']['Reg']), slave=1)
    client.close()
    return ret


def read_v_gain(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=Calibration['V Gain float32']['Start(Dec)'],
                                            count=Calibration['V Gain float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_v_offset(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=Calibration['V offset float32']['Start(Dec)'],
                                            count=Calibration['V offset float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_i_gain(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=Calibration['I Gain float32']['Start(Dec)'],
                                            count=Calibration['I Gain float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def read_i_offset(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=Calibration['I offset float32']['Start(Dec)'],
                                            count=Calibration['I offset float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def get_real_time(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    year = client.read_measurement(address=basic_configuration['Week word']['Start(Dec)'],
                                   count=7, slave=1)
    real_time_clock = str(year[1]) + '-' + str(year[2]).zfill(2) + '-' + str(year[3]).zfill(2) + '-' + str(
        year[0]) + ' ' + str(year[4]).zfill(2) + ':' + str(year[5]).zfill(2) + ':' + str(year[6]).zfill(2)
    client.close()
    return real_time_clock


def read_v_measurement(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    voltage: list = client.read_measurement(address=dc_para_addr['Current float32']['Start(Dec)'],
                                            count=dc_para_addr['Current float32']['Reg'], slave=1)
    client.close()
    reg = hex(voltage[0]) + hex(voltage[1])
    hex_num = reg.replace('0x', '')
    integer_num = int(hex_num, 16)
    return struct.unpack('!f', struct.pack('!I', integer_num))[0]


def get_timestamp(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    year = client.read_measurement(address=16696,
                                   count=7, slave=1)
    real_time_clock = str(year[1]) + '-' + str(year[2]).zfill(2) + '-' + str(year[3]).zfill(2) + '-' + str(
        year[0]) + ' ' + str(year[4]).zfill(2) + ':' + str(year[5]).zfill(2) + ':' + str(year[6]).zfill(2)
    client.close()
    return real_time_clock


def get_usedrecords_datalog1(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=24834, count=2, slave=1)
    client.close()
    return ret


def get_usedrecords_datalog2(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=25090, count=2, slave=1)
    client.close()
    return ret


def get_usedrecords_datalog3(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=25346, count=2, slave=1)
    client.close()
    return ret


def get_usedrecords_datalog4(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.read_measurement(address=25858, count=2, slave=1)
    client.close()
    return ret
