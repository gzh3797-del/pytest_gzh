"""
test_name:
test_num:
author:
modify:
"""
import xlwt
from comm.modbus_get_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_meter_conf_data_get():
    i = 1
    j = 1
    k = 1
    data = {
        'Meter password': [get_meter_password(conn_mode=modbus_config["conn_mode"]), ''],
        'RS485 baud rate': [get_rs485_baudrate(conn_mode=modbus_config["conn_mode"]), 'bps'],
        'RS485 parity': [get_rs485_parity(conn_mode=modbus_config["conn_mode"]), ''],
        'DHCP enable': [get_dhcp_enable(conn_mode=modbus_config["conn_mode"]), ''],
        'IP address': [get_ip_address(conn_mode=modbus_config["conn_mode"]), ''],
        'Subnet mask': [get_subnet_mask(conn_mode=modbus_config["conn_mode"]), ''],
        'Gateway': [get_gateway(conn_mode=modbus_config["conn_mode"]), ''],
        'DNS primary server': [get_dns_primary_server(conn_mode=modbus_config["conn_mode"]), ''],
        'DNS secondary server': [get_dns_secondary_server(conn_mode=modbus_config["conn_mode"]), ''],
        'Modbus Slave ID': [get_slave_id(conn_mode=modbus_config["conn_mode"]), ''],
        'Modbus RTU Enable': [get_rtu_enable(conn_mode=modbus_config["conn_mode"]), ''],
        'Modbus TCP Enable': [get_tcp_enable(conn_mode=modbus_config["conn_mode"]), ''],
        'Modbus TCP port': [get_tcp_port(conn_mode=modbus_config["conn_mode"]), ''],
        'PT1': [get_pt1(conn_mode=modbus_config["conn_mode"]), ''],
        'PT2': [get_pt2(conn_mode=modbus_config["conn_mode"]), ''],
        'CT1': [get_ct1(conn_mode=modbus_config["conn_mode"]), 'A'],
        'CT2': [get_ct2(conn_mode=modbus_config["conn_mode"]), 'mV'],
        'Demand calculation method': [get_demand_calculation_method(conn_mode=modbus_config["conn_mode"]), ''],
        'Demand window time': [get_demand_window_time(conn_mode=modbus_config["conn_mode"]), 'Min'],
        'Demand update period': [get_demand_update_period(conn_mode=modbus_config["conn_mode"]), 'Min'],
        'Energy pulse parameter': [get_energy_pulse_para(conn_mode=modbus_config["conn_mode"]), ''],
        'Energy pulse constant': [get_energy_pulse_constant(conn_mode=modbus_config["conn_mode"]), 'imp/kWh'],
        'Backlight time': [get_backlight_time(conn_mode=modbus_config["conn_mode"]), 'Min'],
        'Seal status': [get_seal_status(conn_mode=modbus_config["conn_mode"]), ''],
        'Device run time': [get_device_run_time(conn_mode=modbus_config["conn_mode"]), 'H'],
        'Device load time': [get_device_load_time(conn_mode=modbus_config["conn_mode"]), 'H'],
        'Enable cable loss compensation': [get_enable_cable_loss_compensation(conn_mode=modbus_config["conn_mode"]), ''],
        'Cable  resistance': [get_cable_resistance(conn_mode=modbus_config["conn_mode"]), 'Ohm']
    }
    system_infomation_data = {
        'Firmware version': [get_firmware_version(conn_mode=modbus_config["conn_mode"]), 'ver'],
        'Firmware release date': [get_firmware_release_date(conn_mode=modbus_config["conn_mode"]), ''],
        'Firmware patch number': [get_firmware_patch_number(conn_mode=modbus_config["conn_mode"]), ''],
        'Serial Number': [get_serial_number(conn_mode=modbus_config["conn_mode"]), ''],
        'Hardware version': [get_hardware_version(conn_mode=modbus_config["conn_mode"]), 'ver'],
        'Function Model Type': [get_function_model_type(conn_mode=modbus_config["conn_mode"]), ''],
        'Voltage Input Type': [get_voltage_input_type(conn_mode=modbus_config["conn_mode"]), ''],
        'Current Input Type': [get_current_input_type(conn_mode=modbus_config["conn_mode"]), ''],
        'Power Supply Type': [get_power_supply_type(conn_mode=modbus_config["conn_mode"]), ''],
        'MAC address': [get_mac_address(conn_mode=modbus_config["conn_mode"]), ''],
        'Device Type': [get_device_type(conn_mode=modbus_config["conn_mode"]), ''],
        'Reserved': [get_reserved(conn_mode=modbus_config["conn_mode"]), ''],
    }
    calibration_data = {
        'V Gain': [read_v_gain(conn_mode=modbus_config["conn_mode"]), ''],
        'V offset': [read_v_offset(conn_mode=modbus_config["conn_mode"]), ''],
        'I Gain': [read_i_gain(conn_mode=modbus_config["conn_mode"]), ''],
        'I offset': [read_i_offset(conn_mode=modbus_config["conn_mode"]), ''],
    }
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Basic Configuration')
    sheet.write(0, 0, 'Descrption')
    sheet.write(0, 1, 'value')
    sheet.write(0, 2, 'Unit')
    for key, value in data.items():
        sheet.write(i, 0, key)
        sheet.write(i, 1, str(value[0]))
        sheet.write(i, 2, str(value[1]))
        i += 1
    sheet = my_workbook.add_sheet('System Infomation')
    sheet.write(0, 0, 'Descrption')
    sheet.write(0, 1, 'value')
    sheet.write(0, 2, 'Unit')
    for key, value in system_infomation_data.items():
        sheet.write(j, 0, key)
        sheet.write(j, 1, str(value[0]))
        sheet.write(j, 2, str(value[1]))
        j += 1
    sheet = my_workbook.add_sheet('Calibration')
    sheet.write(0, 0, 'Descrption')
    sheet.write(0, 1, 'value')
    sheet.write(0, 2, 'Unit')
    for key, value in calibration_data.items():
        sheet.write(k, 0, key)
        sheet.write(k, 1, str(value[0]))
        sheet.write(k, 2, str(value[1]))
        k += 1
    my_workbook.save('{}.xlsx'.format(str(__file__).split("\\")[-1].split('.')[0]))


def teardown_function():
    pass
