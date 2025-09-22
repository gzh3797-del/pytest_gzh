#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @File     :Echilog_Transaction_Test.py
# @Author   :lcs
# @Time     :2025/9/4
# @Desc     :
import struct
import time

from comm.source_control import *
from modbus_config import modbus_config
from comm.modbus_rtu_tcp import ModbusRtuOrTcp

mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def set_OCMF_Command(values):
    """

    :param values:
    0:  do nothing
    0x42:"B" = start transaction
    0x45:"E" = End of transaction
    "A" = Abort transaction
    :return:
    """

    mes.write_registers(address=20992, values=values, slave=1)


def set_enable_cable_loss_compensation(values):
    """

    :param values:0: Disable 1: Enable
    :return:
    """
    mes.write_registers(address=4130, values=values, slave=1)


def read_transaction_import_export_energy():
    energy_list = mes.read_measurement(address=20999, count=8, slave=1)
    transaction_import_energy = hex(energy_list[0]).replace('0x', '').zfill(4) + hex(energy_list[1]).replace('0x',
                                                                                                             '').zfill(
        4) + hex(energy_list[2]).replace('0x', '').zfill(4) + hex(energy_list[3]).replace('0x', '').zfill(4)
    integer_num = int(transaction_import_energy, 16)
    transaction_import_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    transaction_export_energy = hex(energy_list[4]).replace('0x', '').zfill(4) + hex(energy_list[5]).replace('0x',
                                                                                                             '').zfill(
        4) + hex(energy_list[6]).replace('0x', '').zfill(4) + hex(energy_list[7]).replace('0x', '').zfill(4)
    integer_num = int(transaction_export_energy, 16)
    transaction_export_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return transaction_import_energy, transaction_export_energy


def read_device_import_export_energy():
    energy_list = mes.read_measurement(address=16384, count=8, slave=1)
    device_import_energy = hex(energy_list[0]).replace('0x', '').zfill(4) + hex(energy_list[1]).replace('0x', '').zfill(
        4) + hex(energy_list[2]).replace('0x', '').zfill(4) + hex(energy_list[3]).replace('0x', '').zfill(4)
    integer_num = int(device_import_energy, 16)
    device_import_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    device_export_energy = hex(energy_list[4]).replace('0x', '').zfill(4) + hex(energy_list[5]).replace('0x', '').zfill(
        4) + hex(energy_list[6]).replace('0x', '').zfill(4) + hex(energy_list[7]).replace('0x', '').zfill(4)
    integer_num = int(device_export_energy, 16)
    device_export_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    return device_import_energy, device_export_energy


# print(sour_output(voltage=50, current=10))
# sour_stop()


def transaction_log():
    """
    transaction_log快速写满250000条
    :return:
    """
    for i in range(250000):
        time.sleep(0.02)
        set_OCMF_Command(0x42)
        time.sleep(0.02)
        set_OCMF_Command(0x45)
        print(i)


def echilog():
    """快速写满6000条
    echilog
    :return:
    """
    for i in range(6000):
        time.sleep(0.05)
        set_enable_cable_loss_compensation(0)
        time.sleep(0.05)
        set_enable_cable_loss_compensation(1)
        print(i)


def transaction_import_export_energy():
    """
    交易开始时和结束时能量记录与时间记录
    :return:
    """
    sour_output(voltage=200, current=10)
    set_OCMF_Command(0x42)
    print(time.strftime('%Y_%m_%d %H:%M:%S'))
    print(read_device_import_export_energy())
    time.sleep(360)
    set_OCMF_Command(0x45)
    sour_stop()
    print(time.strftime('%Y_%m_%d %H:%M:%S'))
    print(read_device_import_export_energy())


echilog()
# sour_stop()
