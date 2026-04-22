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
import math
import statistics
import struct
import time
from datetime import datetime

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
# from comm.source_control import SourCon
from test_case.AcuRev1320.fast_test.memory_addrs import MemoryAddr, SlaveId, MemoryReg
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
            avg_val_accuracy = round(((avg_val - standard_value) / standard_value), 5)
            min_val_accuracy = round(((min_val - standard_value) / standard_value), 5)
            max_val_accuracy = round(((max_val - standard_value) / standard_value), 5)
        else:
            avg_val_accuracy = round((avg_val - standard_value), 5)
            min_val_accuracy = round((min_val - standard_value), 5)
            max_val_accuracy = round((max_val - standard_value), 5)
            # if all(val == 0 for val in [avg_val, min_val, max_val, standard_value]) or avg_val < 0.001:
            #     (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (0, 0, 0)
            # else:
            #     (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (1, 1, 1)
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

        avg_val_accuracy = round((avg_val - standard_value), 5)
        min_val_accuracy = round((min_val - standard_value), 5)
        max_val_accuracy = round((max_val - standard_value), 5)
        min_val = (min_val + 360) if min_val < 0 else min_val
        max_val = (max_val + 360) if max_val < 0 else max_val
        avg_val = (avg_val + 360) if avg_val < 0 else avg_val
        # if standard_value:
        #     avg_val_accuracy = round((avg_val - standard_value), 5)
        #     min_val_accuracy = round((min_val - standard_value), 5)
        #     max_val_accuracy = round((max_val - standard_value), 5)
        #     min_val = (min_val + 360) if min_val < 0 else min_val
        #     max_val = (max_val + 360) if max_val < 0 else max_val
        #     avg_val = (avg_val + 360) if avg_val < 0 else avg_val
        # else:
        #     if all(val == 0 for val in [avg_val, min_val, max_val, standard_value]) or avg_val < 0.001:
        #         (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (0, 0, 0)
        #     else:
        #         (avg_val_accuracy, min_val_accuracy, max_val_accuracy) = (1, 1, 1)
        return [min_val, min_val_accuracy], [max_val, max_val_accuracy], [avg_val, avg_val_accuracy]

    def read_frequency(self):
        """
        获取寄存器值:频率
        :return: 频率
        """
        address = MemoryAddr.freq_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read frequency ret is:{measure_value}')
        return value

    def read_sys_set_frequency(self):
        """
        获取寄存器值:频率
        :return: 频率
        """
        address = MemoryAddr.set_freq_rms_addr
        count = MemoryReg.reg_single
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!H', bytes(bytes_value))[0]
        self.log.info(f'read sys_set_frequency ret is:{measure_value}')
        return value

    def read_ua_voltage(self):
        """
        获取寄存器值:ua
        :return: ua
        """
        address = MemoryAddr.ua_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        # if type(measure_value) == "str":
        #     measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ua_voltage ret is:{measure_value}')
        return value

    def read_ub_voltage(self):
        """
        获取寄存器值:ub
        :return: ub
        """
        address = MemoryAddr.ub_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ub_voltage ret is:{measure_value}')
        return value

    def read_uc_voltage(self):
        """
        获取寄存器值:uc
        :return: uc
        """
        address = MemoryAddr.uc_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uc_voltage ret is:{measure_value}')
        return value

    def read_uv_avg_voltage(self):
        """
        获取寄存器值:uv_avg
        :return: uv_avg
        """
        address = MemoryAddr.uv_avg_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uv_avg_voltage ret is:{measure_value}')
        return value

    def read_uab_voltage(self):
        """
        获取寄存器值:uab
        :return: uab
        """
        address = MemoryAddr.uab_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uab_voltage ret is:{measure_value}')
        return value

    def read_ubc_voltage(self):
        """
        获取寄存器值:ubc
        :return: ubc
        """
        address = MemoryAddr.ubc_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ubc_voltage ret is:{measure_value}')
        return value

    def read_uca_voltage(self):
        """
        获取寄存器值:uca
        :return: uca
        """
        address = MemoryAddr.uca_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uca_voltage ret is:{measure_value}')
        return value

    def read_ul_avg_voltage(self):
        """
        获取寄存器值:ul_avg
        :return: ul_avg
        """
        address = MemoryAddr.ul_avg_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ul_avg_voltage ret is:{measure_value}')
        return value

    def read_ia_current(self):
        """
        获取寄存器值:ia
        :return: ia
        """
        address = MemoryAddr.ia_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ia_current ret is:{measure_value}')
        return value

    def read_ib_current(self):
        """
        获取寄存器值:ib
        :return: ib
        """
        address = MemoryAddr.ib_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ib_current ret is:{measure_value}')
        return value

    def read_ic_current(self):
        """
        获取寄存器值:ic
        :return: ic
        """
        address = MemoryAddr.ic_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ic_current ret is:{measure_value}')
        return value

    def read_iv_avg_current(self):
        """
        获取寄存器值:iv_avg
        :return: iv_avg
        """
        address = MemoryAddr.iv_avg_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read iv_avg_current ret is:{measure_value}')
        return value

    def read_in_current(self):
        """
        获取寄存器值:in
        :return: in
        """
        address = MemoryAddr.in_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read in_current ret is:{measure_value}')
        return value

    def read_pa_power(self):
        """
        获取寄存器值:pa
        :return: pa
        """
        address = MemoryAddr.pa_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pa_power ret is:{measure_value}')
        # print(f'read pa_power ret is:{value}')
        return value

    def read_pb_power(self):
        """
        获取寄存器值:pb
        :return: pb
        """
        address = MemoryAddr.pb_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pb_power ret is:{measure_value}')
        # print(f'read pb_power ret is:{value}')
        return value

    def read_pc_power(self):
        """
        获取寄存器值:pc
        :return: pc
        """
        address = MemoryAddr.pc_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pc_power ret is:{measure_value}')
        # print(f'read pc_power ret is:{value}')
        return value

    def read_p_total_power(self):
        """
        获取寄存器值:psys
        :return: psys
        """
        address = MemoryAddr.p_total_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read p_total_power ret is:{measure_value}')
        # print(f'read ps_power ret is:{value}')
        return value

    def read_qa_power(self):
        """
        获取寄存器值:qa
        :return: qa
        """
        address = MemoryAddr.qa_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read qa_power ret is:{measure_value}')
        return value

    def read_qb_power(self):
        """
        获取寄存器值:qb
        :return: qb
        """
        address = MemoryAddr.qb_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read qb_power ret is:{measure_value}')
        return value

    def read_qc_power(self):
        """
        获取寄存器值:qc
        :return: qc
        """
        address = MemoryAddr.qc_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read qc_power ret is:{measure_value}')
        return value

    def read_q_total_power(self):
        """
        获取寄存器值:qsys
        :return: qsys
        """
        address = MemoryAddr.q_total_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read q_total_power ret is:{measure_value}')
        return value

    def read_sa_power(self):
        """
        获取寄存器值:sa
        :return: sa
        """
        address = MemoryAddr.sa_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read sa_power ret is:{measure_value}')
        return value

    def read_sb_power(self):
        """
        获取寄存器值:sb
        :return: sb
        """
        address = MemoryAddr.sb_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read sb_power ret is:{measure_value}')
        return value

    def read_sc_power(self):
        """
        获取寄存器值:sc
        :return: sc
        """
        address = MemoryAddr.sc_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read sc_power ret is:{measure_value}')
        return value

    def read_s_total_power(self):
        """
        获取寄存器值:s_sys
        :return: s_sys
        """
        address = MemoryAddr.s_total_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read s_total_power ret is:{measure_value}')
        return value

    def read_pf_a_factor(self):
        """
        获取寄存器值:pf_a
        :return: pf_a
        """
        address = MemoryAddr.pf_a_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pf_a_factor ret is:{measure_value}')
        return value

    def read_pf_b_factor(self):
        """
        获取寄存器值:pf_b
        :return: pf_b
        """
        address = MemoryAddr.pf_b_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pf_b_factor ret is:{measure_value}')
        return value

    def read_pf_c_factor(self):
        """
        获取寄存器值:pf_c
        :return: pf_c
        """
        address = MemoryAddr.pf_c_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pf_c_factor ret is:{measure_value}')
        return value

    def read_pf_total_factor(self):
        """
        获取寄存器值:pf_sys
        :return: pf_sys
        """
        address = MemoryAddr.pf_total_rms_addr
        count = MemoryReg.reg_double
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read pf_total_power ret is:{measure_value}')
        return value

    def read_ua_phase_angle(self):
        """
        获取寄存器值:ua_phase_angle
        :return: ua_phase_angle
        """
        address = MemoryAddr.ua_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        # value = math.radians(value)
        # print(value)
        self.log.info(f'read ua_phase_angle ret is:{measure_value}')
        return value
        # measure_value = [0]
        # value = measure_value[0] / 10
        # self.log.info(f'read ua_phase_angle ret is:{measure_value}')
        # return value

    def read_ub_phase_angle(self):
        """
        获取寄存器值:ub_phase_angle
        :return: ub_phase_angle
        """
        address = MemoryAddr.ub_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        # value = measure_value[0] / 10
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ub_phase_angle ret is:{measure_value}')
        return value

    def read_uc_phase_angle(self):
        """
        获取寄存器值:uc_phase_angle
        :return: uc_phase_angle
        """
        address = MemoryAddr.uc_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        # value = measure_value[0] / 10
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uc_phase_angle ret is:{measure_value}')
        return value

    def read_uab_phase_angle(self):
        """
        获取寄存器值:ua_phase_angle
        :return: ua_phase_angle
        """
        address = MemoryAddr.uab_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        # value = math.radians(value)
        # print(value)
        self.log.info(f'read uab_phase_angle ret is:{measure_value}')
        return value

    def read_ubc_phase_angle(self):
        """
        获取寄存器值:ub_phase_angle
        :return: ub_phase_angle
        """
        address = MemoryAddr.ubc_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ubc_phase_angle ret is:{measure_value}')
        return value

    def read_uca_phase_angle(self):
        """
        获取寄存器值:uc_phase_angle
        :return: uc_phase_angle
        """
        address = MemoryAddr.uca_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read uca_phase_angle ret is:{measure_value}')
        return value

    def read_ia_phase_angle(self):
        """
        获取寄存器值:ia_phase_angle
        :return: ia_phase_angle
        """
        address = MemoryAddr.ia_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ia_phase_angle ret is:{measure_value}')
        return value

    def read_ib_phase_angle(self):
        """
        获取寄存器值:ib_phase_angle
        :return: ib_phase_angle
        """
        address = MemoryAddr.ib_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ib_phase_angle ret is:{measure_value}')
        return value

    def read_ic_phase_angle(self):
        """
        获取寄存器值:ic_phase_angle
        :return: ic_phase_angle
        """
        address = MemoryAddr.ic_phase_angle_rms_addr
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read ic_phase_angle ret is:{measure_value}')
        return value

    def compare_res_by_set_voltage_wire_mode(self, exp_val, act_val):
        """
        判断电压接线方式是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_wire_mode_by_voltage pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_wire_mode_by_voltage fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def compare_res_by_set_current_wire_mode(self, exp_val, act_val):
        """
        判断电流接线方式是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_wire_mode_by_current pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_wire_mode_by_current fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def compare_res(self, exp_val, act_val):
        """
        判断电流接线方式是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_para pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_para fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def compare_res_by_set_clear_energy_flag(self, exp_val, act_val):
        """
        判断清除能量是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_wire_mode_by_voltage pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_wire_mode_by_voltage fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def compare_res_by_set_value(self, exp_val, act_val):
        """
        判断清除能量是否成功设置
        :param exp_val: 期望值
        :param act_val: 寄存器值
        :return: 判断结果
        """
        if exp_val == act_val:
            self.log.info(f'set_value pass, act_val is:{act_val}')
            return True
        else:
            self.log.info(f'set_value fail, exp_val is:{exp_val}, act_val is:{act_val}')
            return False

    def set_wire_mode_by_voltage(self, voltage_wire_mode=4):
        """
        设置电压接线方式
        :param voltage_wire_mode: 电压接线方式
        :return:
        """
        address = MemoryAddr.voltage_wire_addr
        values = voltage_wire_mode
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_wire_mode_by_voltage fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = voltage_wire_mode
        act_val = measure_value[0]
        self.compare_res_by_set_voltage_wire_mode(exp_val, act_val)

    def set_wire_mode_by_current(self, current_wire_mode=0):
        """
        设置电流接线方式
        :param current_wire_mode: 电流接线方式
        :return:
        """
        address = MemoryAddr.current_wire_addr
        values = current_wire_mode
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_wire_mode_by_current fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = current_wire_mode
        act_val = measure_value[0]
        self.compare_res_by_set_current_wire_mode(exp_val, act_val)

    def close_rtu_client(self):
        """
        关闭rtu客户端连接
        :return:
        """
        self.modbus_client.close()

    def read_pa_imp_energy(self):
        """
        获取寄存器值:pa_imp_energy
        :return: pa_imp_energy
        """
        address = MemoryAddr.pa_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pa_imp_energy ret is:{measure_value}')
        return value

    def read_pa_exp_energy(self):
        """
        获取寄存器值:pa_exp_energy_addr
        :return: pa_exp_energy_addr
        """
        address = MemoryAddr.pa_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pa_exp_energy ret is:{measure_value}')
        return value

    def read_pb_imp_energy(self):
        """
        获取寄存器值:pb_imp_energy_addr
        :return: pb_imp_energy_addr
        """
        address = MemoryAddr.pb_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pb_imp_energy ret is:{measure_value}')
        return value

    def read_pb_exp_energy(self):
        """
        获取寄存器值:pb_exp_energy_addr
        :return:pb_exp_energy_addr
        """
        address = MemoryAddr.pb_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pb_exp_energy ret is:{measure_value}')
        return value

    def read_pc_imp_energy(self):
        """
        获取寄存器值:pc_imp_energy_addr
        :return:pc_imp_energy_addr
        """
        address = MemoryAddr.pb_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pc_imp_energy ret is:{measure_value}')
        return value

    def read_pc_exp_energy(self):
        """
        获取寄存器值:pc_exp_energy_addr
        :return:pc_exp_energy_addr
        """
        address = MemoryAddr.pc_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read pc_exp_energy ret is:{measure_value}')
        return value

    def read_qa_imp_energy(self):
        """
        获取寄存器值:qa_imp_energy_addr
        :return:qa_imp_energy_addr
        """
        address = MemoryAddr.qa_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qa_imp_energy ret is:{measure_value}')
        return value

    def read_qa_exp_energy(self):
        """
        获取寄存器值:qa_exp_energy_addr
        :return:qa_exp_energy_addr
        """
        address = MemoryAddr.qa_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qa_exp_energy ret is:{measure_value}')
        return value

    def read_qb_imp_energy(self):
        """
        获取寄存器值:qb_imp_energy_addr
        :return:qb_imp_energy_addr
        """
        address = MemoryAddr.qb_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qb_imp_energy ret is:{measure_value}')
        return value

    def read_qb_exp_energy(self):
        """
        获取寄存器值:qb_exp_energy_addr
        :return:qb_exp_energy_addr
        """
        address = MemoryAddr.qb_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qb_exp_energy ret is:{measure_value}')
        return value

    def read_qc_imp_energy(self):
        """
        获取寄存器值:qc_imp_energy_addr
        :return:qc_imp_energy_addr
        """
        address = MemoryAddr.qc_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qc_imp_energy ret is:{measure_value}')
        return value

    def read_qc_exp_energy(self):
        """
        获取寄存器值:qc_exp_energy_addr
        :return: qc_exp_energy_addr
        """
        address = MemoryAddr.qc_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read qc_exp_energy ret is:{measure_value}')
        return value

    def read_sa_imp_energy(self):
        """
        获取寄存器值:sa_imp_energy_addr
        :return: sa_imp_energy_addr
        """
        address = MemoryAddr.sa_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sa_imp_energy ret is:{measure_value}')
        return value

    def read_sa_exp_energy(self):
        """
        获取寄存器值:sa_exp_energy_addr
        :return: sa_exp_energy_addr
        """
        address = MemoryAddr.sa_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sa_exp_energy ret is:{measure_value}')
        return value

    def read_sb_imp_energy(self):
        """
        获取寄存器值:sb_imp_energy_addr
        :return: sb_imp_energy_addr
        """
        address = MemoryAddr.sb_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sb_imp_energy ret is:{measure_value}')
        return value

    def read_sb_exp_energy(self):
        """
        获取寄存器值:sb_exp_energy_addr
        :return: sb_exp_energy_addr
        """
        address = MemoryAddr.sb_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sb_exp_energy ret is:{measure_value}')
        return value

    def read_sc_imp_energy(self):
        """
        获取寄存器值:sc_imp_energy_addr
        :return: sc_imp_energy_addr
        """
        address = MemoryAddr.sc_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sc_imp_energy ret is:{measure_value}')
        return value

    def read_sc_exp_energy(self):
        """
        获取寄存器值:sc_exp_energy_addr
        :return: sc_exp_energy_addr
        """
        address = MemoryAddr.sc_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sc_exp_energy ret is:{measure_value}')
        return value

    def read_sa_app_energy(self):
        """
        获取寄存器值:sa_app_energy_addr
        :return: sa_app_energy_addr
        """
        address = MemoryAddr.sa_app_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sa_app_energy ret is:{measure_value}')
        return value

    def read_sb_app_energy(self):
        """
        获取寄存器值:sb_app_energy_addr
        :return: sb_app_energy_addr
        """
        address = MemoryAddr.sb_app_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sb_app_energy ret is:{measure_value}')
        return value

    def read_sc_app_energy(self):
        """
        获取寄存器值:sc_app_energy_addr
        :return: sc_app_energy_addr
        """
        address = MemoryAddr.sc_app_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read sc_app_energy ret is:{measure_value}')
        return value

    def read_acc_start_time_energy(self):
        """
        获取寄存器值:acc_start_time_energy_addr
        :return:acc_start_time_energy_addr
        """
        address = MemoryAddr.acc_start_time_energy_addr
        count = MemoryReg.reg_uint16
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!6H', bytes(bytes_value))
        self.log.info(f'read acc_start_time_energy ret is:{measure_value}')
        return value

    def read_acc_end_time_energy(self):
        """
        获取寄存器值:acc_end_time_energy_addr
        :return:acc_end_time_energy_addr
        """
        address = MemoryAddr.acc_end_time_energy_addr
        count = MemoryReg.reg_uint16
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!6H', bytes(bytes_value))[0]
        self.log.info(f'read acc_end_time_energy ret is:{measure_value}')
        return value

    def read_p_sys_imp_energy(self):
        """
        获取寄存器值:p_sys_imp_energy_addr
        :return:p_sys_imp_energy_addr
        """
        address = MemoryAddr.p_sys_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read p_sys_imp_energy ret is:{measure_value}')
        return value

    def read_p_sys_exp_energy(self):
        """
        获取寄存器值:p_sys_exp_energy_addr(
        :return:p_sys_exp_energy_addr(
        """
        address = MemoryAddr.p_sys_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read p_sys_exp_energy_addr( ret is:{measure_value}')
        return value

    def read_p_sys_total_energy(self):
        """
        获取寄存器值:p_sys_total_energy_addr
        :return:p_sys_total_energy_addr
        """
        address = MemoryAddr.p_sys_total_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read p_sys_total_energy ret is:{measure_value}')
        return value

    def read_p_sys_net_energy(self):
        """
        获取寄存器值:p_sys_net_energy_addr
        :return:p_sys_net_energy_addr
        """
        address = MemoryAddr.p_sys_net_energy_addr
        count = MemoryReg.reg_int32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!i', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read p_sys_net_energy ret is:{measure_value}')
        return value

    def read_q_sys_imp_energy(self):
        """
        获取寄存器值:q_sys_imp_energy_addr
        :return:q_sys_imp_energy_addr
        """
        address = MemoryAddr.q_sys_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read q_sys_imp_energy ret is:{measure_value}')
        return value

    def read_q_sys_exp_energy(self):
        """
        获取寄存器值:q_sys_exp_energy_addr
        :return:q_sys_exp_energy_addr
        """
        address = MemoryAddr.q_sys_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read q_sys_exp_energy ret is:{measure_value}')
        return value

    def read_q_sys_total_energy(self):
        """
        获取寄存器值:q_sys_total_energy_addr
        :return:q_sys_total_energy_addr
        """
        address = MemoryAddr.q_sys_total_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read q_sys_total_energy ret is:{measure_value}')
        return value

    def read_q_sys_net_energy(self):
        """
        获取寄存器值:q_sys_net_energy_addr
        :return:q_sys_net_energy_addr
        """
        address = MemoryAddr.q_sys_net_energy_addr
        count = MemoryReg.reg_int32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!i', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read q_sys_net_energy ret is:{measure_value}')
        return value

    def read_s_sys_imp_energy(self):
        """
        获取寄存器值:s_sys_imp_energy_addr
        :return:s_sys_imp_energy_addr
        """
        address = MemoryAddr.s_sys_imp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read s_sys_imp_energy ret is:{measure_value}')
        return value

    def read_s_sys_exp_energy(self):
        """
        获取寄存器值:s_sys_exp_energy_addr
        :return:s_sys_exp_energy_addr
        """
        address = MemoryAddr.s_sys_exp_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read s_sys_exp_energy ret is:{measure_value}')
        return value

    def read_s_sys_total_energy(self):
        """
        获取寄存器值:s_sys_total_energy_addr
        :return:s_sys_total_energy_addr
        """
        address = MemoryAddr.s_sys_total_energy_addr
        count = MemoryReg.reg_uint32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!I', bytes(bytes_value))[0]
        value = value / 1000
        self.log.info(f'read s_sys_total_energy ret is:{measure_value}')
        return value

    def set_cleared_energy(self, clear_energy_flag):
        """
        设置清除能量
        :param clear_energy_flag: 清除能量标志
        :return:
        """
        address = MemoryAddr.clear_energy_addr
        values = clear_energy_flag
        slave = self.slave_id
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_cleared_energy Failed, ret is:{ret}')
            return False
        self.log.info('clear_energy Passed')
        return 'clear_energy Passed'

    def hold_rs485_connect(self, hold_time, time_interval=3):
        """
        每3分钟读取一次 Modbus 数据
        :param hold_time: 保持时间
        :param time_interval: 时间间隔
        :return:
        """
        read_cnt, _ = divmod(hold_time * 60, time_interval)  # 计算读取次数
        for i in range(read_cnt):
            time.sleep(time_interval * 60)
            value = self.read_frequency()
            self.log.info(f"第 {i + 1} 次读取数据:{value},RS485 连接正常")
            print(f"第 {i + 1} 次读取数据:{value},RS485 连接正常")

    def set_demand_method(self, demand_method):
        """
        设置需量算法
        :param demand_method: Fixed Window: 0  Sliding Window: 1
        :return:
        """
        address = MemoryAddr.demand_algorithm_addr
        values = demand_method
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_cleared_energy fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = demand_method
        act_val = measure_value[0]
        self.compare_res_by_set_value(exp_val, act_val)

    def set_demand_interval(self, demand_interval):
        """
        设置上报间隔
        :param demand_interval: 1~30 minute
        :return:
        """
        address = MemoryAddr.demand_interval_addr
        values = demand_interval
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_cleared_energy fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = demand_interval
        act_val = measure_value[0]
        self.compare_res_by_set_value(exp_val, act_val)

    def set_demand_update_rate(self, demand_update_rate):
        """
        设置上报间隔
        :param demand_update_rate: 1~30 minute
        :return:
        """
        address = MemoryAddr.demand_update_rate_addr
        values = demand_update_rate
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_cleared_energy fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = demand_update_rate
        act_val = measure_value[0]
        self.compare_res_by_set_value(exp_val, act_val)

    def read_demand_sys_active_power(self):
        """
        获取寄存器值:system_active_power
        :return:system_active_power
        """
        address = MemoryAddr.demand_addr["system_active_power"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand system_active_power ret is:{measure_value}')
        return value

    def read_demand_sys_reactive_power(self):
        """
        获取寄存器值:system_reactive_power
        :return:system_reactive_power
        """
        address = MemoryAddr.demand_addr["system_reactive_power"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand system_reactive_power ret is:{measure_value}')
        return value

    def read_demand_sys_apparent_power(self):
        """
        获取寄存器值:system_apparent_power
        :return:system_apparent_power
        """
        address = MemoryAddr.demand_addr["system_apparent_power"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand system_apparent_power ret is:{measure_value}')
        return value

    def read_demand_ia(self):
        """
        获取寄存器值:phase_a_current
        :return:phase_a_current
        """
        address = MemoryAddr.demand_addr["phase_a_current"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand phase_a_current ret is:{measure_value}')
        return value

    def read_demand_ib(self):
        """
        获取寄存器值:phase_b_current
        :return:phase_b_current
        """
        address = MemoryAddr.demand_addr["phase_b_current"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand phase_b_current ret is:{measure_value}')
        return value

    def read_demand_ic(self):
        """
        获取寄存器值:phase_c_current
        :return:phase_c_current
        """
        address = MemoryAddr.demand_addr["phase_c_current"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand phase_c_current ret is:{measure_value}')
        return value

    def read_demand_in(self):
        """
        获取寄存器值:phase_n_current
        :return:phase_n_current
        """
        address = MemoryAddr.demand_addr["phase_n_current"]
        count = MemoryReg.reg_float32
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        bytes_value = self.get_bytes_value(measure_value)
        value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read demand phase_n_current ret is:{measure_value}')
        return value

    def set_sys_millisecond(self, sys_millisecond):
        """
        设置上报间隔
        :param sys_millisecond: 
        """
        address = MemoryAddr.sys_millisecond
        values = sys_millisecond
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_sys_millisecond fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = sys_millisecond
        act_val = measure_value[0]
        self.compare_res_by_set_value(exp_val, act_val)

    def set_clear_max_demand(self, clear_max_demand):
        """
        # 重置系统最大需量值，重置需量
        """
        address = MemoryAddr.clear_max_demand
        values = clear_max_demand
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_clear_max_demand fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = clear_max_demand
        act_val = measure_value[0]
        self.compare_res_by_set_value(exp_val, act_val)

    def read_active_power_energy(self):
        """
        获取寄存器值:iv_avg
        :return: iv_avg
        """
        address = MemoryAddr.time_stamp_rms_addr
        count = MemoryReg.reg_timestamp_active_energy
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        # bytes_value = self.get_bytes_value(measure_value)
        # value = struct.unpack('!f', bytes(bytes_value))[0]
        self.log.info(f'read measure_value ret is:{measure_value}')
        return measure_value

    def read_sys_time(self):
        """
        获取寄存器值:sys_time
        :return: sys_time
        """
        address = MemoryAddr.sys_time_rms_addr
        count = MemoryReg.reg_uint16_t * 7
        slave = self.slave_id
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        self.log.info(f"read local_time:{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')}")
        print(f"read local_time:{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')}")
        bytes_value = self.get_bytes_value(measure_value)
        date_time_list = struct.unpack('!7H', bytes(bytes_value))
        self.log.info(f'read sys_time:{date_time_list}')
        print(f'read sys_time:{date_time_list}')
        date_time_obj = datetime(*date_time_list[:-1], date_time_list[-1] * 1000)
        formatted_date_time = date_time_obj.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        self.log.info(f'read sys_time ret is:{measure_value}')
        self.log.info(f'read sys_time ret is:{formatted_date_time}')
        print(f'read sys_time ret is:{formatted_date_time}')
        return date_time_list

    def set_sys_freq(self, freq):
        """
        设置电流接线方式
        :param freq: 频率
        :return:
        """

        address = MemoryAddr.set_freq_rms_addr
        values = freq
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_sys_freq fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = freq
        act_val = measure_value[0]
        self.compare_res(exp_val, act_val)

    def set_phase_order(self, phase_order):
        """
        设置电流接线方式
        :param phase_order: 相序 0:ABC,1:ACB
        :return:
        """
        address = MemoryAddr.phase_order_rms_addr
        values = phase_order
        slave = self.slave_id
        count = MemoryReg.reg_single
        ret = self.modbus_client.write_registers(address=address, values=values, slave=slave)
        if address != ret.address:
            self.log.error(f'set_phase_order fail, ret is:{ret}')
            return False
        measure_value = self.modbus_client.read_measurement(address=address, count=count, slave=slave)
        exp_val = phase_order
        act_val = measure_value[0]
        self.compare_res(exp_val, act_val)


if __name__ == '__main__':
    # syslog
    # rm = HandleMemory()
    # rm.read_sys_time()
    # for i in range(4990):
    #     print(f"第{i}次")
    #     rm.set_cleared_energy(1)
    #     time.sleep(0.1)
    # rm.close_rtu_client()

    # auditlog
    rm = HandleMemory()
    for i in range(2118 // 2):
        print(f"第{i}次")
        time.sleep(0.1)
        rm.set_sys_freq(0)
        time.sleep(0.1)
        rm.set_sys_freq(1)
    rm.close_rtu_client()
