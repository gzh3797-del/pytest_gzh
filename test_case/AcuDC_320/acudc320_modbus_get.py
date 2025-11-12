#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:acuvimseries_modbus_get.py
功能描述:读取寄存器操作模块
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import statistics
import struct
import time

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from comm.source_control import SourCon
from acudc320_memory_info import MemoryAddr, SlaveId, MemoryReg
from tools.log import Log
from modbus_config import modbus_config


class HandleMemory:
    def __init__(self, slave_id=SlaveId.slave_id):
        """
        初始化实例
        :param slave_id: 电表标志slave_id
        """
        self.slave_id = slave_id
        self.modbus_client = None
        self.log = None
        self.init_func()

    def init_func(self):
        """
        ModBus连接,log初始化
        :return:
        """
        self.modbus_client = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
        self.log = Log(str(__file__).split("\\")[-1]).logger

    @staticmethod
    def get_bytes_value(memory_value):
        """
        解析寄存器返回值
        :param memory_value: 寄存器返回值
        :return: 整数列表
        """
        bytes_value = []
        for value in memory_value:
            high_byte = (value & 0xff00) >> 8
            low_byte = (value & 0x00ff)
            bytes_value.extend([high_byte, low_byte])
        return bytes_value

    @staticmethod
    def handle_phase_angle(standard_value, phase_angle_values):
        """
        新增角度0判断
        :param standard_value:输入角度值
        :param phase_angle_values: 测量的相位角度列表
        :return: 转换后的角度值
        """
        if standard_value == 0:
            phase_angle_values = [(phase_angle - 360) if 350 <= phase_angle <= 360 else phase_angle
                                  for phase_angle in phase_angle_values]
            # for i in range(len(phase_angle_values)):
            #     if 350 <= phase_angle_values[i] <= 360:
            #         phase_angle_values[i] = phase_angle_values[i] - 360
        return phase_angle_values

    @staticmethod
    def get_measure_accuracy_by_voltage_current_power(standard_value, measure_values):
        """
        获取电压/电流/功率精度计算结果
        :param standard_value: 输入电压/电流/功率值
        :param measure_values: 寄存器测量值
        :return: 电压/电流/功率精度计算结果
        """
        min_val = min(measure_values)
        max_val = max(measure_values)
        avg_val = statistics.mean(measure_values)
        if standard_value:
            avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
            min_val_accuracy = round(abs((min_val - standard_value) / standard_value), 5)
            max_val_accuracy = round(abs((max_val - standard_value) / standard_value), 5)
        else:
            if all(val == 0 for val in [avg_val, min_val, max_val, standard_value]) or avg_val < 0.001:
                (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (0, 0, 0)
            else:
                (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (1, 1, 1)
        return [min_val, min_val_accuracy], [max_val, max_val_accuracy], [avg_val, avg_val_accuracy]

    @staticmethod
    def get_measure_accuracy_by_phase_angle(standard_value, phase_angle_values):
        """
        获取相位角度精度计算结果
        :param standard_value: 输入相位角度值
        :param phase_angle_values: 寄存器测量值
        :return: 相位角度精度计算结果
        """
        phase_angle_values = HandleMemory.handle_phase_angle(standard_value, phase_angle_values)
        min_val = min(phase_angle_values)
        max_val = max(phase_angle_values)
        avg_val = statistics.mean(phase_angle_values)
        if standard_value:
            avg_val_accuracy = round(abs((avg_val - standard_value) / standard_value), 5)
            min_val_accuracy = round(abs((min_val - standard_value) / standard_value), 5)
            max_val_accuracy = round(abs((max_val - standard_value) / standard_value), 5)
            min_val = (min_val + 360) if min_val < 0 else min_val
            max_val = (max_val + 360) if max_val < 0 else max_val
            avg_val = (avg_val + 360) if avg_val < 0 else avg_val
        else:
            if all(val == 0 for val in [avg_val, min_val, max_val, standard_value]) or avg_val < 0.001:
                (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (0, 0, 0)
            else:
                (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (1, 1, 1)
        return [min_val, min_val_accuracy], [max_val, max_val_accuracy], [avg_val, avg_val_accuracy]

    def read_u_voltage(self):
        """
        获取寄存器值:u
        :return: u
        """
        address = MemoryAddr.u_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read u_voltage ret is:{measure_value}')
        return value

    def read_i1_current(self):
        """
        获取寄存器值:i1
        :return: i1
        """
        address = MemoryAddr.i1_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read i1_current ret is:{measure_value}')
        return value

    def read_i2_current(self):
        """
        获取寄存器值:i2
        :return: i2
        """
        address = MemoryAddr.i2_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read i2_current ret is:{measure_value}')
        return value

    def read_i_sum_current(self):
        """
        获取寄存器值:i2
        :return: i2
        """
        address = MemoryAddr.i_sum_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read i_sum_current ret is:{measure_value}')
        return value

    def read_p1_power(self):
        """
        获取寄存器值:p
        :return: p1
        """
        address = MemoryAddr.p1_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read p1_power ret is:{measure_value}')
        return value

    def read_p2_power(self):
        """
        获取寄存器值:p
        :return: p1
        """
        address = MemoryAddr.p2_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read p2_power ret is:{measure_value}')
        return value

    def read_p_sum_power(self):
        """
        获取寄存器值:p
        :return: p1
        """
        address = MemoryAddr.p_sum_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read p_sum_power ret is:{measure_value}')
        return value

    def close_rtu_client(self):
        """
        关闭rtu客户端连接
        :return:
        """
        self.modbus_client.close()


if __name__ == '__main__':
    rm = HandleMemory()
    j = 0.4
    i_list = []
    for i in range(10):
        i1_current = rm.read_i1_current()
        print(i1_current)
        i_list.append(i1_current)
        accuracy = (i1_current - j) / j
        print(i1_current, accuracy)
        time.sleep(0.1)
    print(i_list)
    rm.close_rtu_client()
