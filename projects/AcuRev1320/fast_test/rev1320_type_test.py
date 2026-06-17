#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:
功能描述:能量测量
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""
import datetime
import inspect
import os
import statistics
import struct
import threading
import time
import math
import cmath
import logging
import socket

from projects.AcuRev1320.fast_test.acuvimseries_modbus_get import HandleMemory
from comm.source_control import (switch_device_screen_interface, set_gear_switching_mode, set_ac, set_voltage_gear,
                                 set_current_gear, up_source_ac)
from projects.AcuRev1320.fast_test.memory_addrs import MemoryAddr
from comm.source_control import (switch_device_screen_interface, set_gear_switching_mode, set_ac, set_voltage_gear,
                                 set_current_gear, up_source_ac)
from tools import device_reboot
import serial

dev_logger = logging.Logger("dev_logger")
log_handler = logging.StreamHandler()
log_handler1 = logging.FileHandler(filename="rev1320_type_test.log", mode='w')
log_handler.setLevel(level=logging.DEBUG)
log_formatter = logging.Formatter(fmt="%(asctime)s: %(levelname)s:  %(message)s")
log_handler.setFormatter(fmt=log_formatter)
log_handler1.setFormatter(fmt=log_formatter)
dev_logger.addHandler(log_handler)
dev_logger.addHandler(log_handler1)

Passed = "[Passed]"
Failed = "[Failed]"


# Failed = "\033[91m{}\033[00m".format("[Failed]")


class DeviceUart:
    def __init__(self, ser_com):
        self.ser = None
        self.port = ser_com
        self.baudrate = 19200
        self.bytesize = 8
        self.parity = None
        self.sparity = 'N'
        self.stopbits = 1
        self.timeout = 0.5
        self.logger = dev_logger
        self.lock = threading.Lock()
        self.connect()

    def connect(self):
        self.ser = serial.Serial(
            port=self.port,
            baudrate=self.baudrate,
            bytesize=self.bytesize,
            parity=self.sparity,
            stopbits=self.stopbits,
            timeout=self.timeout
        )
        connect_ret = "Success" if self.ser.is_open else "Failed"
        self.logger.info(f"Uart,port:{self.ser.port},baudrate:{self.ser.baudrate},connect:{connect_ret}")

    def close(self):
        if not self.ser:
            return f"[Serial] 未连接"
        try:
            self.ser.close()
            self.logger.info(f"[Serial] Connection closed [{self.port}]")
        except (TimeoutError, ConnectionRefusedError, self.timeout, OSError) as e:
            self.logger.error(f"[Serial] Connect Fail to [{self.port}] - {e}")
        finally:
            self.logger.info("[Serial] Connect Execute Completed")

    def send(self, send_data):
        """发送数据"""
        with self.lock:
            self.ser.write(bytes(send_data))
            self.logger.info(f"# send -> {send_data}")

    def receive(self):
        """接收数据"""
        with self.lock:
            number = self.ser.inWaiting
            if number:
                receive_data = self.ser.read(1024)
                self.logger.info(f"# receive -> [{receive_data}]")
                return receive_data

    def get_ser_data(self, send_data):
        self.send(send_data)
        time.sleep(0.001)
        # time.sleep(0.01)
        receive_data = self.receive()
        return receive_data


class DeviceTcp:
    def __init__(self, ip="192.168.1.254", port=502, timeout=0.5):
        self.ip = ip
        self.port = port
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.settimeout(timeout)
        self.isSucess = None
        self.logger = dev_logger
        self.lock = threading.Lock()
        self.connect()

    def connect(self):
        try:
            self.socket.connect((self.ip, self.port))
        except TimeoutError:
            self.isSucess = False
            raise Exception("modbus tcp connect fail")
        else:
            logging.info("modbus tcp connect success")
            self.isSucess = True

    def close(self):
        if not self.socket:
            return f"[TCP] 未连接"
        try:
            self.socket.close()
            self.logger.info(f"[TCP] Connection closed [{self.port}]")
        except (TimeoutError, ConnectionRefusedError, self.socket.timeout, OSError) as e:
            self.logger.error(f"[TCP] Connect Fail to [{self.port}] - {e}")
        finally:
            self.logger.info("[TCP] Connect Execute Completed")

    def send(self, send_data):
        """发送数据"""
        with self.lock:
            try:
                data = bytes(send_data)
                self.socket.sendall(data)
                self.logger.info(f"# send -> {data.hex()}")
            except Exception as e:
                self.logger.error(f"TCP send error: {e}")

    def receive(self):
        """接收数据"""
        with self.lock:
            try:
                receive_data = self.socket.recv(1024)
                if receive_data:
                    self.logger.info(f"# receive -> [{receive_data.hex()}]")
                    return receive_data
            except socket.timeout:
                self.logger.error("TCP receive timeout")
            except KeyboardInterrupt:
                self.logger.error("TCP receive keyboard interrupt")

    def get_ser_data(self, send_data):
        self.send(send_data)
        time.sleep(0.001)
        receive_data = self.receive()
        return receive_data


def get_angle_by_pf(input_pf):
    ua_p = 0
    ub_p = 240
    uc_p = 120
    ia_p = 0
    ib_p = 240
    ic_p = 120
    if input_pf == 0.5:
        ia_p = 300
        ib_p = 180
        ic_p = 60
    elif input_pf == 0.8:
        ia_p = 36.87
        ib_p = 276.87
        ic_p = 156.87
    return ua_p, ub_p, uc_p, ia_p, ib_p, ic_p


def get_voltage(input_voltage):
    ua = input_voltage
    ub = input_voltage
    uc = input_voltage
    return ua, ub, uc


def get_current(input_current):
    ia = input_current
    ib = input_current
    ic = input_current
    return ia, ib, ic


def get_freq(input_freq):
    freq = input_freq
    return freq


def open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq):
    switch_device_screen_interface(inter=0x01)  # 切换至交流界面
    time.sleep(1)
    set_gear_switching_mode('00000000')  # 档位切换归零
    set_voltage_gear(uc, ub, ua)
    set_current_gear(ic, ib, ia)
    set_ac(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    time.sleep(20)


def close_cl3021():
    up_source_ac()
    switch_device_screen_interface(inter=0x00)  # 切换至默认界面
    # time.sleep(5)


def get_time_interval(input_time_interval):
    start_time = time.perf_counter()
    next_time = time.perf_counter() + input_time_interval
    while True:
        remain = next_time - time.perf_counter()
        if remain <= 0:
            break
        time.sleep(min(remain, 0.001))  # 最多睡1ms
    diff_time = time.perf_counter() - start_time
    dev_logger.info(f"本地-间隔时间:{diff_time}s")
    return diff_time


def get_active_power_acc_by_ma(input_current, input_pf, input_offset):
    p_acc = 0
    if input_current == 5 and input_pf == 1:
        p_acc = 0.001 + input_offset

    return p_acc


def get_active_power_acc_by_mv(input_current, input_pf, input_offset):
    p_acc = 0
    if input_current == 5:
        if input_pf == 1:
            p_acc = 0.001 + input_offset
        elif input_pf == 0.5:
            pass
        elif input_pf == 0.8:
            pass
    return p_acc


def get_parer_data(measure_value, ct_type):
    stamp = struct.unpack('!Q', bytes(measure_value[4:12]))[0]
    (
        pa_imp, pb_imp, pc_imp, psys_imp,
        pa_exp, pb_exp, pc_exp, psys_exp,
        pa_total, pb_total, pc_total, psys_total,
        pa_net, pb_net, pc_net, psys_net
    ) = struct.unpack('!12I4i', bytes(measure_value[-66:-2]))
    stamp = stamp / 1000
    # pa_imp = pa_imp
    # pb_imp = pb_imp
    # pc_imp = pc_imp
    # psys_imp = psys_imp
    # pa_exp = pa_exp
    # pb_exp = pb_exp
    # pc_exp = pc_exp
    # psys_exp = psys_exp
    # pa_total = pa_total
    # pb_total = pb_total
    # pc_total = pc_total
    # psys_total = psys_total
    # pa_net = pa_net
    # pb_net = pb_net
    # pc_net = pc_net
    # psys_net = psys_net
    if ct_type == 0:
        # mV - 设置50000, 最大变比10000
        pa_imp = pa_imp / 1000 / 10000
        pb_imp = pb_imp / 1000 / 10000
        pc_imp = pc_imp / 1000 / 10000
        psys_imp = psys_imp / 1000 / 10000
        pa_exp = pa_exp / 1000 / 10000
        pb_exp = pb_exp / 1000 / 10000
        pc_exp = pc_exp / 1000 / 10000
        psys_exp = psys_exp / 1000 / 10000
        pa_total = pa_total / 1000 / 10000
        pb_total = pb_total / 1000 / 10000
        pc_total = pc_total / 1000 / 10000
        psys_total = psys_total / 1000 / 10000
        pa_net = pa_net / 1000 / 10000
        pb_net = pb_net / 1000 / 10000
        pc_net = pc_net / 1000 / 10000
        psys_net = psys_net / 1000 / 10000
    elif ct_type == 1:
        # mA - 设置50000, 最大变比2500
        pa_imp = pa_imp / 1000 / 2500
        pb_imp = pb_imp / 1000 / 2500
        pc_imp = pc_imp / 1000 / 2500
        psys_imp = psys_imp / 1000 / 2500
        pa_exp = pa_exp / 1000 / 2500
        pb_exp = pb_exp / 1000 / 2500
        pc_exp = pc_exp / 1000 / 2500
        psys_exp = psys_exp / 1000 / 2500
        pa_total = pa_total / 1000 / 2500
        pb_total = pb_total / 1000 / 2500
        pc_total = pc_total / 1000 / 2500
        psys_total = psys_total / 1000 / 2500
        pa_net = pa_net / 1000 / 2500
        pb_net = pb_net / 1000 / 2500
        pc_net = pc_net / 1000 / 2500
        psys_net = psys_net / 1000 / 2500
    return stamp, pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


def get_parer_data_by_tcp(measure_value, ct_type):
    stamp = struct.unpack('!Q', bytes(measure_value[10:18]))[0]
    (
        pa_imp, pb_imp, pc_imp, psys_imp,
        pa_exp, pb_exp, pc_exp, psys_exp,
        pa_total, pb_total, pc_total, psys_total,
        pa_net, pb_net, pc_net, psys_net
    ) = struct.unpack('!12I4i', bytes(measure_value[-64:]))
    stamp = stamp / 1000
    if ct_type == 0:
        # mV - 设置50000, 最大变比10000
        pa_imp = pa_imp / 1000 / 10000
        pb_imp = pb_imp / 1000 / 10000
        pc_imp = pc_imp / 1000 / 10000
        psys_imp = psys_imp / 1000 / 10000
        pa_exp = pa_exp / 1000 / 10000
        pb_exp = pb_exp / 1000 / 10000
        pc_exp = pc_exp / 1000 / 10000
        psys_exp = psys_exp / 1000 / 10000
        pa_total = pa_total / 1000 / 10000
        pb_total = pb_total / 1000 / 10000
        pc_total = pc_total / 1000 / 10000
        psys_total = psys_total / 1000 / 10000
        pa_net = pa_net / 1000 / 10000
        pb_net = pb_net / 1000 / 10000
        pc_net = pc_net / 1000 / 10000
        psys_net = psys_net / 1000 / 10000
    elif ct_type == 1:
        # mA - 设置50000, 最大变比2500
        pa_imp = pa_imp / 1000 / 2500
        pb_imp = pb_imp / 1000 / 2500
        pc_imp = pc_imp / 1000 / 2500
        psys_imp = psys_imp / 1000 / 2500
        pa_exp = pa_exp / 1000 / 2500
        pb_exp = pb_exp / 1000 / 2500
        pc_exp = pc_exp / 1000 / 2500
        psys_exp = psys_exp / 1000 / 2500
        pa_total = pa_total / 1000 / 2500
        pb_total = pb_total / 1000 / 2500
        pc_total = pc_total / 1000 / 2500
        psys_total = psys_total / 1000 / 2500
        pa_net = pa_net / 1000 / 2500
        pb_net = pb_net / 1000 / 2500
        pc_net = pc_net / 1000 / 2500
        psys_net = psys_net / 1000 / 2500
    return stamp, pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


# SEND_CMD = [0x01, 0x66, 0x90, 0x50, 0x00, 0x92, 0xA5, 0x7E]
# SEND_CMD_TCP = [0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x01, 0x66, 0xC4, 0x7A, 0x00, 0x92]
SEND_CMD = [0x01, 0x66, 0x90, 0x50, 0x00, 0xB2, 0xA4, 0xA6]
SEND_CMD_TCP = [0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x01, 0x66, 0x90, 0x50, 0x00, 0xB2]


def get_p_energy(voltage, current, pf, time_diff):
    p_power = voltage * current * pf
    if p_power:
        p_imp = (p_power * time_diff) / (1000 * 3600)
        p_exp = 0
        p_total = p_imp + p_exp
        p_net = p_imp - p_exp
    else:
        p_imp = 0
        p_exp = (p_power * time_diff) / (1000 * 3600)
        p_total = p_imp + p_exp
        p_net = p_imp - p_exp
    return p_imp, p_exp, p_total, p_net


def get_exp_p_energy(voltage, current, pf, time_diff):
    p_energy = get_p_energy(voltage, current, pf, time_diff)
    pa_imp, pa_exp, pa_total, pa_net = p_energy
    pb_imp, pb_exp, pb_total, pb_net = p_energy
    pc_imp, pc_exp, pc_total, pc_net = p_energy
    psys_imp, psys_exp, psys_total, psys_net = [p_value * 3 for p_value in p_energy]
    return pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


def get_exp_p_energy_by_ia_0(voltage, current, pf, time_diff):
    p_energy = get_p_energy(voltage, current, pf, time_diff)
    pa_energy = (0, 0, 0, 0)
    pa_imp, pa_exp, pa_total, pa_net = pa_energy
    pb_imp, pb_exp, pb_total, pb_net = p_energy
    pc_imp, pc_exp, pc_total, pc_net = p_energy
    psys_imp, psys_exp, psys_total, psys_net = [p_value * 2 for p_value in p_energy]
    return pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


def get_exp_p_energy_by_iab_0(voltage, current, pf, time_diff):
    p_energy = get_p_energy(voltage, current, pf, time_diff)
    pab_energy = (0, 0, 0, 0)
    pa_imp, pa_exp, pa_total, pa_net = pab_energy
    pb_imp, pb_exp, pb_total, pb_net = pab_energy
    pc_imp, pc_exp, pc_total, pc_net = p_energy
    psys_imp, psys_exp, psys_total, psys_net = [p_value * 1 for p_value in p_energy]
    return pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


def get_exp_p_energy_by_iabc_0(voltage, current, pf, time_diff):
    p_energy = get_p_energy(voltage, current, pf, time_diff)
    pabc_energy = (0, 0, 0, 0)
    pa_imp, pa_exp, pa_total, pa_net = pabc_energy
    pb_imp, pb_exp, pb_total, pb_net = pabc_energy
    pc_imp, pc_exp, pc_total, pc_net = pabc_energy
    psys_imp, psys_exp, psys_total, psys_net = [p_value * 0 for p_value in p_energy]
    return pa_imp, pb_imp, pc_imp, psys_imp, pa_exp, pb_exp, pc_exp, psys_exp, pa_total, pb_total, pc_total, psys_total, pa_net, pb_net, pc_net, psys_net


PARA_DISPLAY = {
    0: "pa_imp",
    1: "pb_imp",
    2: "pc_imp",
    3: "psys_imp",
    4: "pa_exp",
    5: "pb_exp",
    6: "pc_exp",
    7: "psys_exp",
    8: "pa_total",
    9: "pb_total",
    10: "pc_total",
    11: "psys_total",
    12: "pa_net",
    13: "pb_net",
    14: "pc_net",
    15: "psys_net",
}


def cmp_err(exp_energy, act_energy, offset_val):
    for i in range(len(exp_energy)):
        if exp_energy[i]:
            # err_val = (exp_energy[i] - act_energy[i]) / exp_energy[i]
            err_val = (act_energy[i] - exp_energy[i]) / exp_energy[i]
            if abs(err_val) <= offset_val:
                dev_logger.info(
                    f"[Passed], {PARA_DISPLAY[i]}, act:{act_energy[i]}kWh, exp:{exp_energy[i]}kWh, err_val:{'{:.5f}%'.format(err_val * 100)}, exp_val:{'{:.2f}%'.format(offset_val * 100)}")
            else:
                dev_logger.info(
                    f"{Failed}, {PARA_DISPLAY[i]},act:{act_energy[i]}kWh, exp:{exp_energy[i]}kWh, err_val:{'{:.5f}%'.format(err_val * 100)}, exp_val:{'{:.2f}%'.format(offset_val * 100)}")
        else:
            # err_val = (exp_energy[i] - act_energy[i])
            err_val = (act_energy[i] - exp_energy[i])
            if abs(err_val) <= offset_val:
                dev_logger.info(
                    f"{Passed}, {PARA_DISPLAY[i]},act:{act_energy[i]}kWh, exp:{exp_energy[i]}kWh, err_val:{'{:.5f}%'.format(err_val * 100)}, exp_val:{'{:.2f}%'.format(offset_val * 100)}")
            else:
                dev_logger.info(
                    f"{Failed}, {PARA_DISPLAY[i]},act:{act_energy[i]}kWh, exp:{exp_energy[i]}kWh, err_val:{'{:.5f}%'.format(err_val * 100)}, exp_val:{'{:.2f}%'.format(offset_val * 100)}")


def output_energy(act_energy):
    for i in range(len(act_energy)):
        if act_energy[i]:
            dev_logger.info(f"{Failed}, {PARA_DISPLAY[i]}, act:{act_energy[i]}kWh, exp:0kWh")
        else:
            dev_logger.info(f"{Passed}, {PARA_DISPLAY[i]}, act:{act_energy[i]}kWh, exp:0kWh")


# 1、电压输入200V，电流20A，使用脚本获取每路有功输出（取样频率1s），取两个点 (间隔5mins取一个点)，基于能量精度浮动满足±0.02%
def meter_constant_1_ma(input_voltage=200, input_current=20, input_pf=1, input_freq=60, input_time_interval=60,
                        input_err_offset=0.0012, com_id="COM3"):
    # 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 200
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_time_interval = input_time_interval
    input_err_offset = 0.0012

    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def meter_constant_1_mv(input_voltage=200, input_current=5, input_pf=1, input_freq=60, input_time_interval=60,
                        input_err_offset=0.0012, com_id="COM3"):
    # 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 200
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_time_interval = input_time_interval
    input_err_offset = 0.0012

    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def meter_constant_2_ma(input_voltage=200, input_current=20, input_pf=1, input_freq=60):
    pf_values = [1, 0.5, 0.8]
    for i in range(len(pf_values)):
        input_voltage = 200
        input_current = 20
        input_pf = pf_values[i]
        input_freq = input_freq
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        close_cl3021()


def meter_constant_2_mv(input_voltage=200, input_current=5, input_pf=1, input_freq=60):
    pf_values = [1, 0.5, 0.8]
    for i in range(len(pf_values)):
        input_voltage = 200
        input_current = 5
        input_pf = pf_values[i]
        input_freq = input_freq

        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        close_cl3021()


# 3、接入负载200V，5A电流，等待5mins，电表掉电，然后断开负载，电表上电，检查电表每路显示有功能量是否满要求(足基于能量精度浮动满足±0.02%)
def meter_constant_3_ma(input_voltage=200, input_current=20, input_pf=1, input_freq=60, input_time_interval=60,
                        input_err_offset=0.0012, com_id="COM3"):
    # 第一次电表加电
    device_reboot.pow_on_device()
    # 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 200
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_time_interval = input_time_interval
    input_err_offset = 0.0012

    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    local_time_interval = get_time_interval(input_time_interval)
    device_reboot.pow_off_device()
    # 关源+关串口
    device_uart.close()
    close_cl3021()

    # 第二次电表加电
    device_reboot.pow_on_device()
    # time.sleep(5)
    # 第二次读取时间戳和寄存器
    device_uart = DeviceUart(ser_com=com_id)
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    device_reboot.pow_off_device()

    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    # diff_stamp = diff_stamp_energy[0]
    diff_stamp = local_time_interval
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def meter_constant_3_mv(input_voltage=200, input_current=5, input_pf=1, input_freq=60, input_time_interval=60,
                        input_err_offset=0.0012, com_id="COM3"):
    # 第一次电表加电
    device_reboot.pow_on_device()
    # 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 200
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_time_interval = input_time_interval
    input_err_offset = 0.0012

    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    local_time_interval = get_time_interval(input_time_interval)
    device_reboot.pow_off_device()
    # 关源+关串口
    device_uart.close()
    close_cl3021()

    # 第二次电表加电
    device_reboot.pow_on_device()
    # 第二次读取时间戳和寄存器
    device_uart = DeviceUart(ser_com=com_id)
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    device_reboot.pow_off_device()

    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    # diff_stamp = diff_stamp_energy[0]
    diff_stamp = local_time_interval
    dev_logger.info(f"[两次本次时间相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_1_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                  input_time_interval=60, input_err_offset=0.0024, com_id="COM3"):
    # 第一次重复性测试
    # 第一次清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 0.2
        input_pf = 1
        input_err_offset = 0.0024
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 1
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 20
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第四次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 24
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_1_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                  input_time_interval=60, input_err_offset=0.0024, com_id="COM3"):
    # 第一次重复性测试
    # 第一次清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)

        input_voltage = 69
        input_current = 0.05
        input_pf = 1
        input_err_offset = 0.0024

        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 0.25
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 5
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第四次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 6
        input_pf = 1
        input_err_offset = 0.0014
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_0p5l_ma(input_voltage=69, input_current=1, input_pf=0.5, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.0029, com_id="COM3"):
    # 第一次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 1
        input_pf = 0.5
        input_err_offset = 0.0029
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 20
        input_pf = 0.5
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 24
        input_pf = 0.5
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_0p5l_mv(input_voltage=69, input_current=0.25, input_pf=0.5, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.0029, com_id="COM3"):
    # 第一次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 0.25
        input_pf = 0.5
        input_err_offset = 0.0029
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 5
        input_pf = 0.5
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 6
        input_pf = 0.5
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_0p8c_ma(input_voltage=69, input_current=1, input_pf=0.8, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.0029, com_id="COM3"):
    # 第一次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 1
        input_pf = 0.8
        input_err_offset = 0.0029
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 20
        input_pf = 0.8
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 24
        input_pf = 0.8
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def repeatability_test_by_pf_0p8c_mv(input_voltage=69, input_current=0.25, input_pf=0.8, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.0029, com_id="COM3"):
    # 第一次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 0.25
        input_pf = 0.8
        input_err_offset = 0.0029
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第二次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 5
        input_pf = 0.8
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第三次重复性测试 + 清除能量
    for i in range(3):
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 69
        input_current = 6
        input_pf = 0.8
        input_err_offset = 0.0019
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


# 时间保持准确度测试
def test_of_time_keeping_accuracy_ma():
    rm = HandleMemory()
    for i in range(100):
        dev_logger.info(f"第{i + 1}次上下电")
        device_reboot.pow_on_device_of_time_keeping()
        device_reboot.pow_off_device()
    device_reboot.pow_on_device()
    rm.read_sys_time()
    rm.close_rtu_client()
    device_reboot.pow_off_device()


def test_of_time_keeping_accuracy_mv():
    rm = HandleMemory()
    for i in range(100):
        dev_logger.info(f"第{i + 1}次上下电")
        device_reboot.pow_on_device_of_time_keeping()
        device_reboot.pow_off_device()
    device_reboot.pow_on_device()
    rm.read_sys_time()
    rm.close_rtu_client()
    device_reboot.pow_off_device()


# 电压变化
def voltage_variation_ma(input_voltage=60, input_current=0.2, input_pf=1, input_freq=60, input_time_interval=60,
                         input_err_offset=0.0025, com_id="COM3"):
    # 第1次测试 +清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 0.2
    input_pf = 1
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第2次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 0.2
    input_pf = 1
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{2}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第3次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 1
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{3}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第4次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 24
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{4}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第5次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 200
    input_current = 1
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{5}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第6次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 200
    input_current = 24
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{6}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第7次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 1
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{7}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第8次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 24
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{8}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def voltage_variation_mv(input_voltage=60, input_current=0.05, input_pf=1, input_freq=60, input_time_interval=60,
                         input_err_offset=0.0025, com_id="COM3"):
    # 第1次测试 +清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 0.05
    input_pf = 1
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第2次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 0.05
    input_pf = 1
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{2}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第3次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 0.25
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{3}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第4次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 60
    input_current = 6
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{4}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第5次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 200
    input_current = 0.25
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{5}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第6次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 200
    input_current = 6
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{6}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第7次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 0.25
    input_pf = 0.5
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{7}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)

    # 第8次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 480
    input_current = 6
    input_pf = 0.5
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{8}次测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


# 相电压中断
def interruption_of_phase_voltage_ma(input_voltage=69, input_current=20, input_pf=1, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.001, com_id="COM3"):
    # 第1次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.001
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第2次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, input_current, input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{2}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_ia_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第3次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, 0, input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{3}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_iab_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第4次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 20
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, 0, 0)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{4}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_iabc_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)



def interruption_of_phase_voltage_mv(input_voltage=69, input_current=5, input_pf=1, input_freq=60,
                                     input_time_interval=60, input_err_offset=0.001, com_id="COM3"):
    # 第1次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.001
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第2次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, input_current, input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{2}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_ia_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第3次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, 0, input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{3}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_iab_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    # 第4次测试
    # rm = HandleMemory()
    # info = rm.set_cleared_energy(clear_energy_flag=1)
    # dev_logger.info(info)
    # rm.close_rtu_client()
    # time.sleep(5)
    input_voltage = 69
    input_current = 5
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0035
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = (0, 0, 0)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{4}次测试, {input_voltage}V-Ia{ia}A-Ib{ib}A-Ic{ic}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy_by_iabc_0(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


FREQ_DICT = {
    50: 0,
    60: 1,
}


# def set_sys_normal_freq_bak(freq_value):
#     rm = HandleMemory()
#     set_freq = 1
#     if 49 <= freq_value <= 51:
#         set_freq = 0
#     elif 58.8 <= freq_value <= 61.2:
#         set_freq = 1
#     rm.set_sys_freq(freq=set_freq)
#     dev_logger.info(f"set_sys_normal_freq: {set_freq}:{freq_value}Hz")
#     rm.close_rtu_client()


def get_set_freq_val(freq_value):
    set_freq = 1
    if 49 <= freq_value <= 51:
        set_freq = 0
    elif 58.8 <= freq_value <= 61.2:
        set_freq = 1
    return set_freq


def set_sys_normal_freq(freq_value):
    rm = HandleMemory()
    exp_freq = get_set_freq_val(freq_value)
    act_freq = rm.read_sys_set_frequency()
    if exp_freq != act_freq:
        rm.set_sys_freq(exp_freq)
        rm.close_rtu_client()
        time.sleep(5)
        device_reboot.pow_off_device()
        device_reboot.pow_on_device()
        return
    rm.close_rtu_client()


PHASE_ORDER_DICT = {
    0: "ABC",
    1: "ACB",
}


def set_sys_phase_sequence(phase_order):
    rm = HandleMemory()
    set_phase_order = 0
    if phase_order == 0:
        set_phase_order = 0
    else:
        set_phase_order = 1
    rm.set_phase_order(phase_order)
    dev_logger.info(f"set_sys_phase_order: {set_phase_order}:{PHASE_ORDER_DICT[set_phase_order]}")
    rm.close_rtu_client()


# 频率变化
def frequency_variation_ma(input_voltage=347, input_current=0.2, input_pf=1, input_freq=60, input_time_interval=60,
                           input_err_offset=0.0025, com_id="COM3"):
    freq_values = [58.8, 60, 61.2, 49, 50, 51]
    for i in range(len(freq_values)):
        # 第1次测试 + 清除能量 + 设置频率
        set_sys_normal_freq(freq_values[i])
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 0.2
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0025
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-1次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

        # 第2次测试 + 清除能量
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 1
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-2次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

        # 第3次测试 + 清除能量
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 24
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-3次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def frequency_variation_mv(input_voltage=347, input_current=0.05, input_pf=1, input_freq=60, input_time_interval=60,
                           input_err_offset=0.0025, com_id="COM3"):
    freq_values = [58.8, 60, 61.2, 49, 50, 51]
    for i in range(len(freq_values)):
        # 第1次测试 + 清除能量 + 设置频率
        set_sys_normal_freq(freq_values[i])
        # if i == 3:
        #     device_reboot.pow_off_device()
        #     time.sleep(5)
        #     device_reboot.pow_on_device()
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 0.05
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0025
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-1次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

        # 第2次测试 + 清除能量
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 0.25
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-2次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)

        # 第3次测试 + 清除能量
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 6
        input_pf = 1
        input_freq = freq_values[i]
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}-3次测试开始, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)


# 相序颠倒
def reversed_phase_sequence_ma(input_voltage=347, input_current=20, input_pf=1, input_freq=60, input_time_interval=60,
                               input_err_offset=0.0015, com_id="COM3"):
    phase_order = [0, 1]
    for i in range(len(phase_order)):
        # 第1次测试 + 清除能量 + 设置频率
        set_sys_normal_freq(input_freq)
        set_sys_phase_sequence(phase_order[i])
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 20
        input_pf = 1
        input_freq = input_freq
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 1)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次-[{phase_order[i]}:{PHASE_ORDER_DICT[phase_order[i]]}]测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    set_sys_phase_sequence(phase_order=0)


def reversed_phase_sequence_mv(input_voltage=347, input_current=5, input_pf=1, input_freq=60, input_time_interval=60,
                               input_err_offset=0.0015, com_id="COM3"):
    phase_order = [0, 1]
    for i in range(len(phase_order)):
        # 第1次测试 + 清除能量 + 设置频率
        set_sys_normal_freq(input_freq)
        set_sys_phase_sequence(phase_order[i])
        rm = HandleMemory()
        info = rm.set_cleared_energy(clear_energy_flag=1)
        dev_logger.info(info)
        rm.close_rtu_client()
        time.sleep(5)
        input_voltage = 347
        input_current = 5
        input_pf = 1
        input_freq = input_freq
        input_err_offset = 0.0015
        ua, ub, uc = get_voltage(input_voltage)
        ia, ib, ic = get_current(input_current)
        ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
        freq = get_freq(input_freq)
        # 开源+开串口
        open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
        device_uart = DeviceUart(ser_com=com_id)
        # 第一次读取时间戳和寄存器
        measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
        get_time_interval(input_time_interval)
        # 第二次读取时间戳和寄存器
        measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
        # 关源+关串口
        close_cl3021()
        device_uart.close()
        # 计算两次之间的时间和间隔时间戳和精度
        stamp_energy_1st = get_parer_data(measure_value_1st, 0)
        stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
        diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
        diff_stamp = diff_stamp_energy[0]
        dev_logger.info(
            f"第{i + 1}次-[{phase_order[i]}-{PHASE_ORDER_DICT[phase_order[i]]}]测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
        dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
        exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
        act_p_energy = diff_stamp_energy[1:]
        cmp_err(exp_p_energy, act_p_energy, input_err_offset)
    set_sys_phase_sequence(phase_order=0)


# 辅助电压变化
def auxiliary_voltage_variation_ac80_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                        input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    """
    测试方法:
    步骤13: 毫安电表(Imin 0.2A，)，供电电压遍历步骤1~15中数值， 数据精度均满足要求。
    备注1: 基于有功能量精度浮动满足±0.04%,  例如测点能量精度为0.2，则当前满足要求0.16% ~ 0.24%
    备注2：供电AC电压范围100V ~ 480V，
    """
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第1次测试, mA型, AC=80V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac100_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第2次测试, mA型, AC=100V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac480_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第3次测试, mA型, AC=480V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac552_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第4次测试, mA型, AC=552V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac80_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                        input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    """
    测试方法:
    步骤1: 毫伏电表，AC供电(最小供电-20%) 80V
    步骤2: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤3: 基于有功能量精度浮动±0.05%
    步骤4: 毫伏电表，AC供电(最小供电-0%) 100V
    步骤5: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤6: 基于有功能量精度浮动±0.05%

    步骤7:毫伏电表，AC供电(最大供电-0%)480V
    步骤8:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤9: 基于有功能量精度浮动±0.05%
    步骤10:毫伏电表，AC供电(最大供电+15%)552V
    步骤11:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤12: 基于有功能量精度浮动±0.05%
    """
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第1次测试, mV型, AC=80V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac100_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    """
    测试方法:
    步骤1: 毫伏电表，AC供电(最小供电-20%) 80V
    步骤2: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤3: 基于有功能量精度浮动±0.05%
    步骤4: 毫伏电表，AC供电(最小供电-0%) 100V
    步骤5: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤6: 基于有功能量精度浮动±0.05%

    步骤7:毫伏电表，AC供电(最大供电-0%)480V
    步骤8:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤9: 基于有功能量精度浮动±0.05%
    步骤10:毫伏电表，AC供电(最大供电+15%)552V
    步骤11:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤12: 基于有功能量精度浮动±0.05%
    """
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第2次测试, mV型, AC=100V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac480_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    """
    测试方法:
    步骤1: 毫伏电表，AC供电(最小供电-20%) 80V
    步骤2: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤3: 基于有功能量精度浮动±0.05%
    步骤4: 毫伏电表，AC供电(最小供电-0%) 100V
    步骤5: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤6: 基于有功能量精度浮动±0.05%

    步骤7:毫伏电表，AC供电(最大供电-0%)480V
    步骤8:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤9: 基于有功能量精度浮动±0.05%
    步骤10:毫伏电表，AC供电(最大供电+15%)552V
    步骤11:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤12: 基于有功能量精度浮动±0.05%
    """
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第3次测试, mV型, AC=480V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def auxiliary_voltage_variation_ac552_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                         input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    """
    测试方法:
    步骤1: 毫伏电表，AC供电(最小供电-20%) 80V
    步骤2: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤3: 基于有功能量精度浮动±0.05%
    步骤4: 毫伏电表，AC供电(最小供电-0%) 100V
    步骤5: 接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤6: 基于有功能量精度浮动±0.05%

    步骤7:毫伏电表，AC供电(最大供电-0%)480V
    步骤8:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤9: 基于有功能量精度浮动±0.05%
    步骤10:毫伏电表，AC供电(最大供电+15%)552V
    步骤11:接入电流Imin 0.05A，最小额定电压69V，PF为1
    步骤12: 基于有功能量精度浮动±0.05%
    """
    set_sys_normal_freq(input_freq)
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第4次测试, mV型, AC=552V, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


# 辅助装置的运行
def operation_of_auxiliary_devices_rtu_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                          input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    # 第1次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次-ma-rtu-测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def operation_of_auxiliary_devices_rtu_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                          input_time_interval=60, input_err_offset=0.0025, com_id="COM3"):
    # 第1次测试
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_uart = DeviceUart(ser_com=com_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_uart.get_ser_data(send_data=SEND_CMD)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_uart.get_ser_data(send_data=SEND_CMD)
    # 关源+关串口
    close_cl3021()
    device_uart.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次-[mv]-[rtu]-测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def operation_of_auxiliary_devices_tcp_ma(input_voltage=69, input_current=0.2, input_pf=1, input_freq=60,
                                          input_time_interval=60, input_err_offset=0.0025,
                                          tcp_id=("192.168.1.254", 502)):
    # 第1次测试 + 清除能量
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.2
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_tcp = DeviceTcp(*tcp_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_tcp.get_ser_data(send_data=SEND_CMD_TCP)
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_tcp.get_ser_data(send_data=SEND_CMD_TCP)
    # 关源+关串口
    close_cl3021()
    device_tcp.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data_by_tcp(measure_value_1st, 1)
    stamp_energy_2nd = get_parer_data_by_tcp(measure_value_2nd, 1)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次-ma-tcp-测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


def operation_of_auxiliary_devices_tcp_mv(input_voltage=69, input_current=0.05, input_pf=1, input_freq=60,
                                          input_time_interval=60, input_err_offset=0.0025,
                                          tcp_id=("192.168.1.254", 502)):
    # 第1次测试
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)
    input_voltage = 69
    input_current = 0.05
    input_pf = 1
    input_freq = input_freq
    input_err_offset = 0.0025
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 开源+开串口
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_tcp = DeviceTcp(*tcp_id)
    # 第一次读取时间戳和寄存器
    measure_value_1st = device_tcp.get_ser_data(send_data=SEND_CMD_TCP)
    dev_logger.info(list(measure_value_1st))
    get_time_interval(input_time_interval)
    # 第二次读取时间戳和寄存器
    measure_value_2nd = device_tcp.get_ser_data(send_data=SEND_CMD_TCP)
    dev_logger.info(list(measure_value_2nd))
    # 关源+关串口
    close_cl3021()
    device_tcp.close()
    # 计算两次之间的时间和间隔时间戳和精度
    stamp_energy_1st = get_parer_data_by_tcp(measure_value_1st, 0)
    stamp_energy_2nd = get_parer_data_by_tcp(measure_value_2nd, 0)
    diff_stamp_energy = [_2nd - _1st for _2nd, _1st in zip(stamp_energy_2nd, stamp_energy_1st)]
    diff_stamp = diff_stamp_energy[0]
    dev_logger.info(
        f"第{1}次-mv-tcp-测试, {input_voltage}V-{input_current}A-pf_{input_pf}-{input_freq}Hz-{input_err_offset * 100}%")
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{diff_stamp}s")
    exp_p_energy = get_exp_p_energy(input_voltage, input_current, input_pf, diff_stamp)
    act_p_energy = diff_stamp_energy[1:]
    cmp_err(exp_p_energy, act_p_energy, input_err_offset)


# 仪表初始启动
def initial_start_up_of_the_meter_ma(input_voltage=69, input_current=24, input_pf=1, input_freq=60,
                                     input_time_interval=10, com_id="COM3"):
    # 表上电-清能量-下电
    device_reboot.pow_on_device()
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    device_reboot.pow_off_device()

    input_voltage = 69
    input_current = 24
    input_pf = 1
    input_freq = input_freq
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 加源-表上电-间隔10s-下电-关源
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    start_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    dev_logger.info(f"电表上电时间: {start_time}")
    device_reboot.pow_on_device()
    get_time_interval(input_time_interval)
    device_reboot.pow_off_device()
    close_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    dev_logger.info(f"电表下电时间: {close_time}")
    close_cl3021()
    time.sleep(5)
    device_reboot.pow_on_device()
    time.sleep(5)
    device_uart = DeviceUart(ser_com=com_id)
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{input_time_interval}s")
    measure_value = device_uart.get_ser_data(send_data=SEND_CMD)
    stamp_energy = get_parer_data(measure_value, 1)
    output_energy(stamp_energy[1:])
    device_reboot.pow_off_device()


def initial_start_up_of_the_meter_mv(input_voltage=69, input_current=6, input_pf=1, input_freq=60,
                                     input_time_interval=10, com_id="COM3"):
    # 表上电-清能量-下电
    device_reboot.pow_on_device()
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    device_reboot.pow_off_device()

    input_voltage = 69
    input_current = 6
    input_pf = 1
    input_freq = input_freq
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 加源-表上电-间隔10s-下电-关源
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_reboot.pow_on_device()
    get_time_interval(input_time_interval)
    device_reboot.pow_off_device()
    close_cl3021()
    device_reboot.pow_on_device()
    device_uart = DeviceUart(ser_com=com_id)
    dev_logger.info(f"[两次寄存器值相减-时间间隔], 两次读取时间间隔的能量累积时间:{input_time_interval}s")
    measure_value = device_uart.get_ser_data(send_data=SEND_CMD)
    stamp_energy = get_parer_data(measure_value, 0)
    output_energy(stamp_energy[1:])
    device_reboot.pow_off_device()


# 启动电流测试
def starting_current_test_ma(input_voltage=69, input_current=0.012, input_pf=1, input_freq=60):
    # 表上电-清能量-下电
    device_reboot.pow_on_device()
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 69
    input_current = 0.012
    input_pf = 1
    input_freq = input_freq
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 加源-表上电-间隔10s-下电-关源
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_reboot.pow_off_device()
    close_cl3021()


def starting_current_test_mv(input_voltage=69, input_current=0.00375, input_pf=1, input_freq=60):
    # 表上电-清能量-下电
    device_reboot.pow_on_device()
    rm = HandleMemory()
    info = rm.set_cleared_energy(clear_energy_flag=1)
    dev_logger.info(info)
    rm.close_rtu_client()
    time.sleep(5)

    input_voltage = 69
    input_current = 0.00375
    input_pf = 1
    input_freq = input_freq
    ua, ub, uc = get_voltage(input_voltage)
    ia, ib, ic = get_current(input_current)
    ua_p, ub_p, uc_p, ia_p, ib_p, ic_p = get_angle_by_pf(input_pf)
    freq = get_freq(input_freq)
    # 加源-表上电-间隔10s-下电-关源
    open_cl3021(uc_p, ub_p, ua_p, ic_p, ib_p, ia_p, uc, ub, ua, ic, ib, ia, freq)
    device_reboot.pow_off_device()
    close_cl3021()


def modify_freq(cnt):
    rm = HandleMemory()
    for i in range(cnt):
        print(f"第{i}次")
        rm.set_sys_freq(1)
        time.sleep(0.1)
        rm.set_sys_freq(0)
        time.sleep(0.1)
    rm.close_rtu_client()


# def is_device_reboot_by_freq(self, cfg_freq):
#     freq_dict = {
#         50: 0,
#         60: 1,
#     }
#     exp_freq = freq_dict[cfg_freq]
#     act_freq = self.handle_memory.read_sys_set_frequency()
#     if exp_freq != act_freq:
#         self.handle_memory.set_sys_freq(exp_freq)
#         self.handle_memory.close_rtu_client()
#         device_reboot.pow_off_device()
#         time.sleep(5)
#         device_reboot.pow_on_device()
#         self.handle_memory = HandleMemory(slave_id=1)


if __name__ == '__main__':
    # mA
    # 01 电表常数
    device_reboot.pow_on_device()
    # meter_constant_1_ma()
    # meter_constant_2_ma()
    # meter_constant_3_ma()
    # device_reboot.pow_off_device()
    # 02 仪表初始启动
    # initial_start_up_of_the_meter_ma()
    # 04 启动电流测试
    # device_reboot.pow_on_device()
    # starting_current_test_ma()
    # device_reboot.pow_off_device()
    # 05 重复性测试
    # device_reboot.pow_on_device()
    # repeatability_test_by_pf_1_ma()
    # repeatability_test_by_pf_0p5l_ma()
    # repeatability_test_by_pf_0p8c_ma()
    # device_reboot.pow_off_device()
    # 07 时间保持准确度测试
    # device_reboot.pow_on_device()
    # test_of_time_keeping_accuracy_ma()
    # device_reboot.pow_off_device()
    # 08 电压变化
    # device_reboot.pow_on_device()
    # voltage_variation_ma()
    # device_reboot.pow_off_device()
    # 09 相电压中断
    # device_reboot.pow_on_device()
    # interruption_of_phase_voltage_ma()
    # device_reboot.pow_off_device()
    # 10 频率变化
    # device_reboot.pow_on_device()
    # frequency_variation_ma()
    # device_reboot.pow_off_device()
    # 11 相序颠倒
    # device_reboot.pow_on_device()
    # reversed_phase_sequence_ma()
    # device_reboot.pow_off_device()
    # 12 辅助电压变化
    # device_reboot.pow_on_device()
    # auxiliary_voltage_variation_ac80_ma()
    # auxiliary_voltage_variation_ac100_ma()
    # auxiliary_voltage_variation_ac480_ma()
    # auxiliary_voltage_variation_ac552_ma()
    # device_reboot.pow_off_device()
    # 13 辅助装置的运行
    # device_reboot.pow_on_device()
    # operation_of_auxiliary_devices_rtu_ma()
    # operation_of_auxiliary_devices_tcp_ma()
    # device_reboot.pow_off_device()

    # mV
    # 01 电表常数
    # device_reboot.pow_on_device()
    # meter_constant_1_mv()
    # meter_constant_2_mv()
    # meter_constant_3_mv()
    # device_reboot.pow_off_device()
    # 02 仪表初始启动
    # initial_start_up_of_the_meter_mv()
    # 04 启动电流测试
    # device_reboot.pow_on_device()
    # starting_current_test_mv()
    # device_reboot.pow_off_device()
    # 05 重复性测试
    # device_reboot.pow_on_device()
    # repeatability_test_by_pf_1_mv()
    # repeatability_test_by_pf_0p5l_mv()
    # repeatability_test_by_pf_0p8c_mv()
    # device_reboot.pow_off_device()
    # 07 时间保持准确度测试
    # device_reboot.pow_on_device()
    # test_of_time_keeping_accuracy_mv()
    # device_reboot.pow_off_device()
    # 08 电压变化
    # device_reboot.pow_on_device()
    # voltage_variation_mv()
    # device_reboot.pow_off_device()
    # 09 相电压中断
    # device_reboot.pow_on_device()
    # interruption_of_phase_voltage_mv()
    # device_reboot.pow_off_device()
    # 10 频率变化
    # device_reboot.pow_on_device()
    # frequency_variation_mv()
    # device_reboot.pow_off_device()
    # 11 相序颠倒
    # device_reboot.pow_on_device()
    # reversed_phase_sequence_mv()
    # device_reboot.pow_off_device()
    # 12 辅助电压变化
    # device_reboot.pow_on_device()
    # auxiliary_voltage_variation_ac80_mv()
    # auxiliary_voltage_variation_ac100_mv()
    # auxiliary_voltage_variation_ac480_mv()
    # auxiliary_voltage_variation_ac552_mv()
    # device_reboot.pow_off_device()
    # 13 辅助装置的运行
    # device_reboot.pow_on_device()
    # operation_of_auxiliary_devices_rtu_mv()
    # operation_of_auxiliary_devices_tcp_mv()
    # device_reboot.pow_off_device()
