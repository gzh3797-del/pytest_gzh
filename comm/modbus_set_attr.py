import time
from comm.modbus_get_attr import *
from modbus_config import write_json, modbus_config


def set_meter_password(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Meter password word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4096,1)' not in str(ret):
        logging.error('meter password set fail, ret is:{}'.format(ret))
        return False
    ret = get_meter_password(conn_mode=conn_mode)
    if ret != value:
        logging.error('meter password ret is:{}'.format(ret))
        return False
    return True


def set_pt1(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['PT1 word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4114,1)' in str(ret):
        logging.error('set nonsupport set, but pt1 set success, ret is:{}'.format(ret))
        return True
    return False


def set_pt2(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['PT2 word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4115,1)' in str(ret):
        logging.error('set nonsupport set, but pt2 set success, ret is:{}'.format(ret))
        return True
    return False


def set_ct1(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['CT1 word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4116,1)' in str(ret):
        logging.error('set nonsupport set, but CT1 set success, ret is:{}'.format(ret))
        return True
    return False


def set_ct2(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['CT2 word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4117,1)' in str(ret):
        logging.error('set nonsupport set, but CT2 set success, ret is:{}'.format(ret))
        return True
    return False


def set_seal_status(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Seal status word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4125,1)' in str(ret):
        logging.error('set nonsupport set, but Seal status set success, ret is:{}'.format(ret))
        return True
    logging.info('Seal status set ret is:{}'.format(ret))
    return False


def set_cable_loss_compensation_enable(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Enable cable loss compensation word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4130,1)' not in str(ret):
        logging.error('set cable loss compensation enable ret is:{}'.format(ret))
        return False
    ret = get_enable_cable_loss_compensation(conn_mode=conn_mode)
    if ret != value:
        logging.info('cable loss compensation enable ret is:{}'.format(ret[0]))
        return False
    return True


def set_cable_resistance(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Cable  resistance word']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(4131,1)' not in str(ret):
        logging.error('set cable resistance ret is:{}'.format(ret))
        return False
    ret = get_cable_resistance(conn_mode=conn_mode)
    if ret != value:
        logging.info('cable resistance ret is:{}'.format(ret))
        return False
    return True


def clear_energy(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear energy word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8192,1)' not in str(ret):
        logging.error('clear energy fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear energy word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear energy word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def clear_charge(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear charge word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8193,1)' not in str(ret):
        logging.error('clear charge fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear charge word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear charge word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def clear_demand(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear demand word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8194,1)' not in str(ret):
        logging.error('clear demand fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear demand word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear demand word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def clear_max_min(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear max/min word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8195,1)' not in str(ret):
        logging.error('clear max/min fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear max/min word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear max/min word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def clear_device_run_time(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear device run time word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8196,1)' not in str(ret):
        logging.error('clear device run tim fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear device run time word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear device run time word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def clear_load_time(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear load time word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8197,1)' not in str(ret):
        logging.error('clear load time fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear load time word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear load time word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def factory_reset(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Factory reset word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8198,1)' not in str(ret):
        logging.error('factory reset fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Factory reset word']['Start(Dec)'],
                                  count=int(basic_configuration['Factory reset word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def network_reset(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Clear charge word']['Start(Dec)'],
                                 values=[value], slave=1)
    if '(8193,1)' not in str(ret):
        logging.error('network reset fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=basic_configuration['Clear charge word']['Start(Dec)'],
                                  count=int(basic_configuration['Clear charge word']['Reg']), slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def set_rs485_baudrate(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['RS485 baud rate word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4097,1)' not in str(ret):
        logging.error('rs485 baudrate set fail, ret is:{}'.format(ret))
        return False
    if value == 7680:
        write_json('baudrate', 76800)
    elif value == 11520:
        write_json('baudrate', 115200)
    elif value == 12800:
        write_json('baudrate', 128000)
    else:
        write_json('baudrate', value)
    time.sleep(1)
    ret = get_rs485_baudrate(conn_mode=conn_mode)
    if ret != value:
        logging.error('rs485 baudrate ret is:{}'.format(ret))
        return False
    return True


def set_rs485_parity(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['RS485 parity word']['Start(Dec)'], values=[value],
                                 slave=1)
    time.sleep(1)
    client.close()
    if '(4098,1)' not in str(ret):
        logging.error('rs485 parity set fail, ret is:{}'.format(ret))
        return False
    if value == 0:
        write_json('parity', 'E')
    elif value == 1:
        write_json('parity', 'O')
    else:
        write_json('parity', 'N')
    ret = get_rs485_parity(conn_mode=conn_mode)
    if ret != value:
        logging.error('rs485 parity ret is:{}'.format(ret))
        return False
    return True


def set_dhcp_enable(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['DHCP enable word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4099,1)' not in str(ret):
        logging.error('DHCP enable set fail, ret is:{}'.format(ret))
        return False
    ret = get_dhcp_enable(conn_mode=conn_mode)
    if ret != value:
        logging.error('DHCP enable ret is:{}'.format(ret))
        return False
    return True


def set_ip_address(conn_mode, value):
    ip = value.split('.')
    ip = [hex(int(i)) for i in ip]
    ip1 = int(ip[0][2:4].zfill(2) + ip[1][2:4].zfill(2), 16)
    ip2 = int(ip[2][2:4].zfill(2) + ip[3][2:4].zfill(2), 16)
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(
        address=basic_configuration['IP address 1st byte (high)\nIP address 2nd byte (low) word']['Start(Dec)'],
        values=[ip1, ip2],
        slave=1)

    client.close()
    write_json('ip', value)
    if '(4100,2)' not in str(ret):
        logging.error('IP address set fail, ret is:{}'.format(ret))
        return False
    ret = get_ip_address(conn_mode=conn_mode)
    if ret != value:
        logging.error('IP address ret is:{}'.format(ret))
        return False
    return True


def set_subnet_mask(conn_mode, value):
    sub = value.split('.')
    sub = [hex(int(i)).replace('0x', '').zfill(4) for i in sub]
    sub1 = int(sub[0][2:5] + sub[1][2:5], 16)
    sub2 = int(sub[2][2:5] + sub[3][2:5], 16)
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(
        address=basic_configuration['Subnet mask 1st byte (high)\nSubnet mask 2nd byte (low) word']['Start(Dec)'],
        values=[sub1, sub2],
        slave=1)

    client.close()
    if '(4102,2)' not in str(ret):
        logging.error('Subnet mask set fail, ret is:{}'.format(ret))
        return False
    ret = get_subnet_mask(conn_mode=conn_mode)
    if ret != value:
        logging.error('Subnet mask ret is:{}'.format(ret))
        return False
    return True


def set_gateway(conn_mode, value):
    sub = value.split('.')
    sub = [hex(int(i)).replace('0x', '').zfill(4) for i in sub]
    sub1 = int(sub[0][2:5] + sub[1][2:5], 16)
    sub2 = int(sub[2][2:5] + sub[3][2:5], 16)
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(
        address=basic_configuration['Gateway 1st byte (high)\nGateway 2nd byte (low) word']['Start(Dec)'],
        values=[sub1, sub2],
        slave=1)

    client.close()
    if '(4104,2)' not in str(ret):
        logging.error('Gateway set fail, ret is:{}'.format(ret))
        return False
    ret = get_gateway(conn_mode=conn_mode)
    if ret != value:
        logging.error('Gateway ret is:{}'.format(ret))
        return False
    return True


def set_dns_primary(conn_mode, value):
    sub = value.split('.')
    sub = [hex(int(i)).replace('0x', '').zfill(4) for i in sub]
    sub1 = int(sub[0][2:5] + sub[1][2:5], 16)
    sub2 = int(sub[2][2:5] + sub[3][2:5], 16)
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(
        address=basic_configuration['DNS primary server 1st byte (high)\nDNS primary server 2nd byte (low) word'][
            'Start(Dec)'],
        values=[sub1, sub2],
        slave=1)

    client.close()
    if '(4106,2)' not in str(ret):
        logging.error('DNS primary server set fail, ret is:{}'.format(ret))
        return False
    ret = get_dns_primary_server(conn_mode=conn_mode)
    if ret != value:
        logging.error('DNS primary server ret is:{}'.format(ret))
        return False
    return True


def set_dns_secondary(conn_mode, value):
    sub = value.split('.')
    sub = [hex(int(i)).replace('0x', '').zfill(4) for i in sub]
    sub1 = int(sub[0][2:5] + sub[1][2:5], 16)
    sub2 = int(sub[2][2:5] + sub[3][2:5], 16)
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(
        address=basic_configuration['DNS secondary server 1st byte (high)\nDNS secondary server 2nd byte (low) word'][
            'Start(Dec)'],
        values=[sub1, sub2],
        slave=1)

    client.close()
    if '(4108,2)' not in str(ret):
        logging.error('DNS secondary server set fail, ret is:{}'.format(ret))
        return False
    ret = get_dns_secondary_server(conn_mode=conn_mode)
    if ret != value:
        logging.error('DNS secondary server ret is:{}'.format(ret))
        return False
    return True


def set_modbus_slaveid(conn_mode, value):
    slave = get_slave_id(conn_mode=conn_mode, slave=modbus_config['rtu']['slaveid'])
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Modbus Slave ID word']['Start(Dec)'], values=[value],
                                 slave=slave)
    client.close()
    if '(4110,1)' not in str(ret):
        logging.error('Modbus Slave ID set fail, ret is:{}'.format(ret))
        return False
    write_json('slaveid', value)
    ret = get_slave_id(conn_mode=conn_mode, slave=value)
    if ret != value:
        logging.error('Modbus Slave ID ret is:{}'.format(ret))
        return False
    return True


def set_modbus_rtu_enable(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Modbus RTU Enable word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4111,1)' not in str(ret):
        logging.error('Modbus RTU Enable set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_rtu_enable(conn_mode=conn_mode)
    if ret != value:
        logging.error('Modbus RTU Enable ret is:{}'.format(ret))
        return False
    return True


def set_modbus_tcp_enable(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Modbus TCP Enable word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4112,1)' not in str(ret):
        logging.error('Modbus TCP Enable set fail, ret is:{}'.format(ret))
        return False
    ret = get_tcp_enable(conn_mode=conn_mode)
    if ret != value:
        logging.error('Modbus TCP Enable ret is:{}'.format(ret))
        return False
    return True


def set_modbus_tcp_port(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Modbus TCP port word']['Start(Dec)'], values=[value],
                                 slave=1)
    client.close()
    if '(4113,1)' not in str(ret):
        logging.error('Modbus TCP port set fail, ret is:{}'.format(ret))
        return False
    write_json('port', value)
    # time.sleep(3)
    ret = get_tcp_port(conn_mode=conn_mode)
    if ret != value:
        logging.error('Modbus TCP port ret is:{}'.format(ret))
        return False
    return True


def set_demand_calculation_method(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Demand calculation method word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4118,1)' not in str(ret):
        logging.error('Demand calculation method set fail, ret is:{}'.format(ret))
        return False
    ret = get_demand_calculation_method(conn_mode=conn_mode)
    if ret != value:
        logging.error('Demand calculation method ret is:{}'.format(ret))
        return False
    return True


def set_demand_window_time(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Demand window time word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4119,1)' not in str(ret):
        logging.error('Demand window time set fail, ret is:{}'.format(ret))
        return False
    ret = get_demand_window_time(conn_mode=conn_mode)
    if ret != value:
        logging.error('Demand window time ret is:{}'.format(ret))
        return False
    return True


def set_demand_update_period(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Demand update period word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4120,1)' not in str(ret):
        logging.error('Demand update period set fail, ret is:{}'.format(ret))
        return False
    ret = get_demand_update_period(conn_mode=conn_mode)
    if ret != value:
        logging.error('Demand update period ret is:{}'.format(ret))
        return False
    return True


def set_energy_pulse_parameter(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Energy pulse parameter word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4121,1)' not in str(ret):
        logging.error('Energy pulse parameter set fail, ret is:{}'.format(ret))
        return False
    ret = get_energy_pulse_para(conn_mode=conn_mode)
    if ret != value:
        logging.error('Energy pulse parameter ret is:{}'.format(ret))
        return False
    return True


def set_energy_pulse_constant(conn_mode, value):
    value *= 1000
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Energy pulse constant word']['Start(Dec)'],
                                 values=[int(value) // 65536, int(value) % 65536],
                                 slave=1)
    client.close()
    if '(4122,2)' not in str(ret):
        logging.error('Energy pulse constant set fail, ret is:{}'.format(ret))
        return False
    ret = get_energy_pulse_constant(conn_mode=conn_mode)
    if ret != round(value / 1000, 3):
        logging.error('Energy pulse constant ret is:{}'.format(ret))
        return False
    return True


def set_backlight_time(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Backlight time word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4124,1)' not in str(ret):
        logging.error('Backlight time set fail, ret is:{}'.format(ret))
        return False
    ret = get_backlight_time(conn_mode=conn_mode)
    if ret != value:
        logging.error('Backlight time constant ret is:{}'.format(ret))
        return False
    return True


def set_device_run_time(conn_mode, value):
    time_h = value / 3600
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Device run time (high 16 bits) word']['Start(Dec)'],
                                 values=[int(value) % 65535, int(value) // 65535],
                                 slave=1)
    client.close()
    if '(4126,2)' not in str(ret):
        logging.error('Device run time set fail, ret is:{}'.format(ret))
        return False
    ret = get_device_run_time(conn_mode=conn_mode)
    if ret != round(time_h, 2):
        logging.error('Device run time ret is:{}, set time is:{}'.format(ret, round(time_h, 2)))
        return False
    return True


def set_device_load_time(conn_mode, value):
    time_h = value / 3600
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Device load time (high 16 bits) word']['Start(Dec)'],
                                 values=[int(value) % 65535, int(value) // 65535],
                                 slave=1)
    client.close()
    if '(4128,2)' not in str(ret):
        logging.error('Device load time set fail, ret is:{}'.format(ret))
        return False
    ret = get_device_load_time(conn_mode=conn_mode)
    if ret != round(time_h, 2):
        logging.error('Device load time ret is:{}, set time is:{}'.format(ret, round(time_h, 2)))
        return False
    return True


def set_year(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Year word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4160,1)' not in str(ret):
        logging.error('Year set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[0:4]) != value:
        logging.error('Year constant ret is:{}'.format(int(ret[0:4])))
        return False
    return True


def set_month(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Month word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4161,1)' not in str(ret):
        logging.error('Month set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[5:7]) != value:
        logging.error('Month  ret is:{}'.format(int(ret[5:7])))
        return False
    return True


def set_day(conn_mode, value):
    ret = None
    if conn_mode == 'rtu':
        client = ModbusRtuOrTcp(conn_mode=conn_mode)
        ret = client.write_registers(address=basic_configuration['Day word']['Start(Dec)'],
                                     values=[value],
                                     slave=1)
        client.close()
    if conn_mode == 'tcp':
        client = ModbusRtuOrTcp(conn_mode=conn_mode)
        ret = client.write_registers(address=basic_configuration['Day word']['Start(Dec)'],
                                     values=[value],
                                     slave=1)
        client.close()
    if '(4162,1)' not in str(ret):
        logging.error('Day set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[8:10]) != value:
        logging.error('Day ret is:{}'.format(int(ret[8:10])))
        return False
    return True


def set_week(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Week word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4159,1)' not in str(ret):
        logging.error('Week set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[11:12]) != value:
        logging.error('Week ret is:{}'.format(int(ret[11:12])))
        return False
    return True


def set_hour(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Hour word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4163,1)' not in str(ret):
        logging.error('Hour set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[13:15]) != value:
        logging.error('Hour ret is:{}'.format(int(ret[13:15])))
        return False
    return True


def set_minute(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Minute word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4164,1)' not in str(ret):
        logging.error('Minute set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if int(ret[16:18]) != value:
        logging.error('Minute ret is:{}'.format(ret))
        return False
    return True


def set_second(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=basic_configuration['Second word']['Start(Dec)'],
                                 values=[value],
                                 slave=1)
    client.close()
    if '(4165,1)' not in str(ret):
        logging.error('Second set fail, ret is:{}'.format(ret))
        return False
    time.sleep(1)
    ret = get_real_time(conn_mode=conn_mode)
    if value == 59:
        if int(ret[19:21]) == 0:
            return True
        else:
            logging.error('Second ret is:{}'.format(ret))
            return False
    if int(ret[19:21]) != value + 1:
        logging.error('Second ret is:{}'.format(int(ret[19:21])))
        return False
    return True


def set_voltage_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Measured) float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12288,1)' not in str(ret):
        logging.error('V(Measured) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_measurement_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Measured) int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12809,1)' not in str(ret):
        logging.error('V(Measured) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_measu_or_comp_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Measured or Compensated) int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12800,1)' not in str(ret):
        logging.error('V(Measured or Compensated) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_current_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Current int16']['Start(Dec)'],
                                 values=value, slave=1)
    client.close()
    if '(12801,1)' not in str(ret):
        logging.error('Current set fail, ret is:{}'.format(ret))
        return False
    return True


def set_power_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Power int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12802,1)' not in str(ret):
        logging.error('Power set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_ripple_factor_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Voltage Ripple Factor int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12803,1)' not in str(ret):
        logging.error('Voltage Ripple Factor set fail, ret is:{}'.format(ret))
        return False
    return True


def set_current_ripple_factor_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Current Ripple Factor int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12804,1)' not in str(ret):
        logging.error('Current Ripple Factor set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_current_import_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Current import int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12805,1)' not in str(ret):
        logging.error('Demand Current import set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_current_export_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Current export int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12806,1)' not in str(ret):
        logging.error('Demand Current export set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_power_import_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Power import int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12807,1)' not in str(ret):
        logging.error('Demand Power import set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_power_export_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Power export int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12808,1)' not in str(ret):
        logging.error('Demand Power export set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_comp_client(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Compensated) int16']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12810,1)' not in str(ret):
        logging.error('V(Compensated) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_measu_or_comp(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Measured or Compensated) float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12288,1)' not in str(ret):
        logging.error('V(Measured or Compensated) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_compensated(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['V(Compensated) float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12308,1)' not in str(ret):
        logging.error('V(Compensated) set fail, ret is:{}'.format(ret))
        return False
    return True


def set_current_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Current float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12290,1)' not in str(ret):
        logging.error('Current set fail, ret is:{}'.format(ret))
        return False
    return True


def set_power_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Power float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12292,1)' not in str(ret):
        logging.error('Power set fail, ret is:{}'.format(ret))
        return False
    return True


def set_voltage_ripple_factor_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Voltage Ripple Factor float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12294,1)' not in str(ret):
        logging.error('Voltage Ripple Factor set fail, ret is:{}'.format(ret))
        return False
    return True


def set_current_ripple_factor_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Current Ripple Factor float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12296,1)' not in str(ret):
        logging.error('Current Ripple Factor set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_current_import_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Current import float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12298,1)' not in str(ret):
        logging.error('Demand Current import set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_current_export_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Current export float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12300,1)' not in str(ret):
        logging.error('Demand Current export set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_power_import_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Power import float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12302,1)' not in str(ret):
        logging.error('Demand Power import set fail, ret is:{}'.format(ret))
        return False
    return True


def set_demand_power_export_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Demand Power export float32']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(12304,1)' not in str(ret):
        logging.error('Demand Power export set fail, ret is:{}'.format(ret))
        return False
    return True


def set_import_energy_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Import Energy double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16384,1)' not in str(ret):
        logging.error('Import Energy set fail, ret is:{}'.format(ret))
        return False
    return True


def set_export_energy_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Export Energy double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16388,1)' not in str(ret):
        logging.error('Export Energy set fail, ret is:{}'.format(ret))
        return False
    return True


def set_net_energy_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Net Energy double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16392,1)' not in str(ret):
        logging.error('Net Energy set fail, ret is:{}'.format(ret))
        return False
    return True


def set_total_energy_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Total Energy double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16396,1)' not in str(ret):
        logging.error('Total Energy set fail, ret is:{}'.format(ret))
        return False
    return True


def set_import_charge_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Import Charge double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16400,1)' not in str(ret):
        logging.error('Import Charge set fail, ret is:{}'.format(ret))
        return False
    return True


def set_export_charge_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Export Charge double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16404,1)' not in str(ret):
        logging.error('Export Charge set fail, ret is:{}'.format(ret))
        return False
    return True


def set_net_charge_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Net Charge double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16408,1)' not in str(ret):
        logging.error('Net Charge set fail, ret is:{}'.format(ret))
        return False
    return True


def set_total_charge_measurement(conn_mode, value):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=dc_para_addr['Total Charge double']['Start(Dec)'],
                                 values=[value], slave=1)
    client.close()
    if '(16412,1)' not in str(ret):
        logging.error('Net Charge set fail, ret is:{}'.format(ret))
        return False
    return True


def clear_data_log1(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=4411,
                                 values=1, slave=1)
    client.close()
    return ret


def clear_data_log2(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=4471,
                                 values=1, slave=1)
    client.close()
    return ret


def clear_data_log3(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=4531,
                                 values=1, slave=1)
    client.close()
    return ret


def clear_data_log4(conn_mode):
    client = ModbusRtuOrTcp(conn_mode=conn_mode)
    ret = client.write_registers(address=5435,
                                 values=1, slave=1)
    client.close()
    return ret



def set_identification_status(value):
    """
    set identification status
    :param value: identification status
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5000, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x5000, count=1, slave=1)
    client.close()
    if '(20480,1)' not in str(ret):
        logging.error('identification status set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != value:
        logging.error('identification status set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_identification_level(value):
    """
    set identification level
    :param value: identification level
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5001, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x5001, count=1, slave=1)
    client.close()
    if '(20481,1)' not in str(ret):
        logging.error('identification level set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != value:
        logging.error('identification level set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_identification_flag(value: list):
    """
    set identification flag
    :param value: 列表长度为4，每一个元素为十六进制数，占一个字节
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5002, values=value, slave=1)
    ret1 = client.read_measurement(address=0x5002, count=4, slave=1)
    client.close()
    if '(20482,4)' not in str(ret):
        logging.error('identification flag set fail, ret is:{}'.format(ret))
        return False
    if list(ret1) != value:
        logging.error('identification flag set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_identification_type(value):
    """
    set identification type
    :param value: identification type
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5006, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x5006, count=1, slave=1)
    client.close()
    if '(20486,1)' not in str(ret):
        logging.error('identification type set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != value:
        logging.error('identification type set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_identification_data(value: list):
    """
    set identification data
    :param value: 列表长度最大为20，每一个元素为十六进制数，占一个字节
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5007, values=value, slave=1)
    ret1 = client.read_measurement(address=0x5007, count=len(value), slave=1)
    client.close()
    if '(20487,{})'.format(len(value)) not in str(ret):
        logging.error('identification type set fail, ret is:{}'.format(ret))
        return False
    if list(ret1) != value:
        logging.error('identification type set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_tariff_text(value: list):
    """
    set tariff text
    :param value: 列表长度最大为20，每一个元素为十六进制数，占一个字节
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x501B, values=value, slave=1)
    ret1 = client.read_measurement(address=0x501B, count=len(value), slave=1)
    client.close()
    if '(20507,{})'.format(len(value)) not in str(ret):
        logging.error('identification type set fail, ret is:{}'.format(ret))
        return False
    if list(ret1) != value:
        logging.error('identification type set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_transaction_overtime(value):
    """
    set transaction overtime
    :param value: transaction overtime
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x502F, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x502F, count=1, slave=1)
    client.close()
    if '(20527,1)' not in str(ret):
        logging.error('transaction overtime set fail, ret is:{}'.format(ret))
        return False
    if ret1 != value:
        logging.error('transaction overtime set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_ct(value):
    """
    set ct
    :param value: ct
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5030, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x5030, count=1, slave=1)
    client.close()
    if '(20528,1)' not in str(ret):
        logging.error('ct set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != value:
        logging.error('ct set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_ci(value: list):
    """
    set tariff text
    :param value: 列表长度最大为20，每一个元素为十六进制数，占一个字节
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5031, values=value, slave=1)
    ret1 = client.read_measurement(address=0x5031, count=len(value), slave=1)
    client.close()
    if '(20529,{})'.format(len(value)) not in str(ret):
        logging.error('ci set fail, ret is:{}'.format(ret))
        return False
    if list(ret1) != value:
        logging.error('ci set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_timezone_shift(value):
    """
    set timezone shift
    :param value: timezone shift
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5045, values=[value], slave=1)
    ret1 = client.read_measurement(address=0x5045, count=1, slave=1)
    client.close()
    if '(20549,1)' not in str(ret):
        logging.error('timezone shift set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != value:
        logging.error('timezone shift set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_utc_timestamp(value):
    """
    set utc timestamp
    :param value: utc timestamp
    :return: True--Set Success, False--Set Fail
    """
    hex_value = hex(value).replace('0x', '').zfill(8)
    client = ModbusRtuOrTcp()
    print([int('0x' + hex_value[0:2], 16), int('0x' + hex_value[2:4], 16),
           int('0x' + hex_value[4:6], 16), int('0x' + hex_value[6:8], 16)])
    ret = client.write_registers(address=0x5046,
                                 values=[int('0x' + hex_value[0:2], 16), int('0x' + hex_value[2:4], 16),
                                         int('0x' + hex_value[4:6], 16), int('0x' + hex_value[6:8], 16)],
                                 slave=1)
    ret1 = client.read_measurement(address=0x5046, count=4, slave=1)
    client.close()
    ret2 = int(hex(ret1[0]).replace('0x', '') + hex(ret1[1]).replace('0x', '') + hex(ret1[2]).replace('0x', '') + hex(
        ret1[3]).replace('0x', ''), 16)
    if '(20550,4)' not in str(ret):
        logging.error('utc timestamp set fail, ret is:{}'.format(ret))
        return False
    if ret2 != value:
        logging.error('utc timestamp shift set fail, ret is:{}'.format(ret2))
        return False
    return True


def clear_transaction_log():
    """
    clear transaction log
    :param value:
    :return: True-- Success, False-- Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x50A1, values=[1], slave=1)
    client.close()
    if '(20641,1)' not in str(ret):
        logging.error('clear transaction log set fail, ret is:{}'.format(ret))
        return False
    return True


def ocmf_command(value: str):
    """
    ocmf command
    :param value: ocmf command
    :return: True--Set Success, False--Set Fail
    """
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5200, values=[ord(value)], slave=1)
    ret1 = client.read_measurement(address=0x5200, count=1, slave=1)
    client.close()
    if '(20992,1)' not in str(ret):
        logging.error('ocmf command set fail, ret is:{}'.format(ret))
        return False
    if ret1[0] != ord(value):
        logging.error('ocmf command set fail, ret is:{}'.format(ret1))
        return False
    return True


def set_transactionlog_id(value):
    """
    set transactionlog id
    :param value: transactionlog id
    :return: True--Set Success, False--Set Fail
    """
    hex_value = hex(value).replace('0x', '').zfill(4)
    client = ModbusRtuOrTcp()
    ret = client.write_registers(address=0x5217,
                                 values=[int('0x' + hex_value[0:2], 16), int('0x' + hex_value[2:4], 16)],
                                 slave=1)
    ret1 = client.read_measurement(address=0x5217, count=4, slave=1)
    client.close()
    ret2 = int(hex(ret1[0]).replace('0x', '') + hex(ret1[1]).replace('0x', ''), 16)
    if '(21015,2)' not in str(ret):
        logging.error('transactionlog id set fail, ret is:{}'.format(ret))
        return False
    if ret2 != value:
        logging.error('transactionlog id set fail, ret is:{}'.format(ret2))
        return False
    return True
