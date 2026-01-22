#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:dev_uart.py
功能描述:
创建日期:2026/1/13
作者:
版本:v1.0
修改记录:
"""
import inspect
import math
import struct
import threading
import time

import serial

import logging

dev_logger = logging.Logger("dev_logger")
log_handler = logging.StreamHandler()
log_handler1 = logging.FileHandler(filename="pt2000.log", mode='w')
log_handler.setLevel(level=logging.DEBUG)
log_formatter = logging.Formatter(fmt="%(asctime)s: %(levelname)s:  %(message)s")
log_handler.setFormatter(fmt=log_formatter)
log_handler1.setFormatter(fmt=log_formatter)
dev_logger.addHandler(log_handler)
dev_logger.addHandler(log_handler1)

FRAME_ID = 0x5A5A
CMD_ONLINE = 0x004C  # 联机命令，读取终端型号和版本号
CMD_SET_FREQ = 0x0071  # 设置标准源输出频率
CMD_SET_WIRE = 0x0072  # 设置接线方式
CMD_SET_GEAR = 0x0073  # 设置标准源挡位
CMD_SET_AMP = 0x0074  # 设置标准源电压电流的幅值
CMD_SET_PHASE = 0x0075  # 设置标准源电压电流的相位
CMD_SET_P_POWER = 0x0076  # 设置有功功率
CMD_SET_Q_POWER = 0x0077  # 设置无功功率
CMD_SET_HARMONIC_PARA = 0x0078  # 设置谐波参数
CMD_ENABLE_SYS_SET_PARA = 0x0079  # 系统控制设置命令
CMD_SET_ENERGY_VERIFY_PARA = 0x007A  # 设置电能校验参数
CMD_ENABLE_ENERGY_VERIFY = 0x007B  # 启动电能校验
CMD_ENABLE_HARMONIC_PARA = 0x007D  # 根据设置的谐波参数输出谐波
# CMD_SET_CLOSE_LOOP = 0x0079  # 设置闭环控制使能

CMD_ONLINE_LEN = 0x0004  # 联机命令，读取终端型号和版本号,指令长度
CMD_SET_FREQ_LEN = 0x0006  # 设置标准源输出频率,指令长度
CMD_SET_WIRE_LEN = 0x0005  # 设置接线方式,指令长度
CMD_SET_GEAR_LEN = 0x000A  # 设置标准源挡位,指令长度
CMD_SET_AMP_LEN = 0x0010  # 设置标准源电压电流的幅值,指令长度
CMD_SET_PHASE_LEN = 0x0010  # 设置标准源电压电流的相位,指令长度
CMD_SET_P_POWER_LEN = 0x0007  # 设置有功功率,指令长度
CMD_SET_Q_POWER_LEN = 0x0007  # 设置无功功率,指令长度
# CMD_SET_HARMONIC_PARA_LEN = 0x0078  # 设置谐波参数,指令长度
# CMD_SET_CLOSE_LOOP_LEN = 0x0079  # 设置闭环控制使能,指令长度
CMD_SET_ENERGY_VERIFY_PARA_LEN = 0x0010  # 设置电能校验参数,指令长度
CMD_ENABLE_ENERGY_VERIFY_LEN = 0x0005  # 启动电能校验,指令长度
CMD_ENABLE_SYS_SET_PARA_LEN = 0x0007  # 系统控制设置命令,指令长度
CMD_ENABLE_HARMONIC_PARA_LEN = 0x0005  # 根据设置的谐波参数输出谐波,指令长度

CMD_SET_SYS_MODE = 0x0080  # 设置系统模式
CMD_SET_SWELL_OR_SWAG = 0x0081  # 设置骤升/骤降
CMD_SET_FLICKER = 0x0082  # 设置闪变
CMD_SET_INTER_HARMONIC = 0x008f  # 设置间谐波
CMD_ENABLE_INTER_HARMONIC = 0x008d  # 根据设置的间谐波参数输出间谐波

CMD_SET_SYS_MODE_LEN = 0x0005  # 设置系统模式,指令长度
CMD_SET_SWELL_OR_SWAG_LEN = 0x004C  # 设置骤升/骤降,指令长度
CMD_SET_FLICKER_LEN = 0x002E  # 设置闪变,指令长度
# CMD_SET_INTER_HARMONIC_LEN = 0x008f  # 设置间谐波,指令长度
CMD_ENABLE_INTER_HARMONIC_LEN = 0x0005  # 根据设置的间谐波参数输出间谐波,指令长度

CMD_GET_METER_PARA = 0x0090  # 读交流标准表参数
CMD_GET_HARMONIC = 0x0091  # 读谐波参数
CMD_GET_3_PHASE_HARMONIC = 0x009B  # 读3相谐波参数
CMD_GET_FLICKER = 0x0092  # 读闪变参数
CMD_GET_3_PHASE_INTER_HARMONIC = 0x009F  # 读3相间谐波参数
CMD_GET_SWELL_OR_SWAG = 0x0094  # 读骤升/骤降参数
CMD_SEND_ERROR_CODE = 0x0095  # 发送错误代码，一旦检测到错误代码，系统自动发送

CMD_GET_METER_PARA_LEN = 0x0004  # 读交流标准表参数,指令长度
# CMD_GET_HARMONIC_LEN = 0x0091  # 读谐波参数,指令长度
CMD_GET_3_PHASE_HARMONIC_LEN = 0x0004  # 读3相谐波参数,指令长度
CMD_GET_FLICKER_LEN = 0x0004  # 读闪变参数,指令长度
CMD_GET_3_PHASE_INTER_HARMONIC_LEN = 0x0004  # 读读3相间谐波参数,指令长度
CMD_GET_SWELL_OR_SWAG_LEN = 0x0004  # 读骤升/骤降参数,指令长度
CMD_SEND_ERROR_CODE_LEN = 0x0005  # 发送错误代码，一旦检测到错误代码，系统自动发送,指令长度

CMD_SWITCH_AMP = 0x0151  # 切换源幅值校准点
CMD_CALIBRATE_AMP = 0x0152  # 校准标准源幅值
CMD_CONFIRM_AMP = 0x0153  # 确认校准标准源幅值
CMD_RESET_FACTORY_SETTING = 0x0156  # 清空校准参数，恢复出厂默认值

CMD_SWITCH_AMP_LEN = 0x0007  # 切换源幅值校准点,指令长度
CMD_CALIBRATE_AMP_LEN = 0x001F  # 校准标准源幅值,指令长度
CMD_CONFIRM_AMP_LEN = 0x0004  # 确认校准标准源幅值,指令长度
CMD_RESET_FACTORY_SETTING_LEN = 0x0007  # 清空校准参数，恢复出厂默认值,指令长度
CHUNKSIZE = 1024


class DeviceUart:
    def __init__(self, ser_com):
        self.ser = None
        self.port = ser_com
        self.baudrate = 115200
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

    def send(self, send_data: bytes):
        """发送数据"""
        with self.lock:
            time.sleep(1)
            self.ser.write(bytes(send_data))
            sd_data_str = ['{:02x}'.format(data) for data in send_data]
            sd_data = " ".join(sd_data_str)
            proto_func_name = inspect.stack()[1].function
            self.logger.info(f"# {proto_func_name} |    send -> [{sd_data}]")

    def receive(self):
        """接收数据"""
        with self.lock:
            time.sleep(1)
            number = self.ser.inWaiting
            if number:
                receive_data = self.ser.read(CHUNKSIZE)
                rcv_data = ['{:02x}'.format(data) for data in receive_data]
                proto_func_name = inspect.stack()[1].function
                _data = " ".join(rcv_data)
                self.logger.info(f"# {proto_func_name} | receive -> [{_data}]")
                return list(receive_data)

    @staticmethod
    def get_parse_data(*datas):
        send_cmd = []
        for data in datas:
            high_bits = (data & 0xFF00) >> 8
            low_bits = (data & 0x00FF)
            send_cmd.extend([high_bits, low_bits])
        return send_cmd

    def get_frame_id(self):
        data = FRAME_ID
        frame_id = self.get_parse_data(data)
        return frame_id

    def get_checksum(self, data_lst):
        # data = sum(data_lst)
        # checksum = self.get_parse_data(data)
        # return checksum

        _data_lst = []
        for i in range(0, len(data_lst), 2):
            _data = (data_lst[i] << 8) | (data_lst[i + 1])
            _data_lst.append(_data)
        data = sum(_data_lst)
        checksum = self.get_parse_data(data)
        return checksum

    @staticmethod
    def parse_bytes_to_ascii_by_unpack(data):
        """解析字节为字符串"""
        ret = bytes(data).decode('utf8')
        return ret
        # bdata = bytes(data)
        # if isinstance(data, list):
        #     bdata = bytes(data)
        # elif isinstance(data, (bytes, bytearray)):
        #     bdata = data
        # fmt = f'{len(bdata)}H'
        # return struct.unpack(fmt, bdata)[0].decode('ascii')

    @staticmethod
    def parse_bytes_to_u32_by_unpack(data):
        """
        解析字节为字符串
        将 4 字节数据解析为 uint32
        endian:
            '>' 大端（Modbus 常见）
            '<' 小端
        """
        if isinstance(data, list):
            bdata = bytes(data)
        elif isinstance(data, (bytes, bytearray)):
            bdata = data
        else:
            raise TypeError("data must be list[int] / bytes / bytearray")
        if len(bdata) != 4:
            raise ValueError("uint32 必须是 4 字节")
        return struct.unpack(f'!I', bdata)[0]

    @staticmethod
    def parse_bytes_to_u16_by_unpack_patch(data_lst):
        """
        解析字节为字符串
        将 2 字节数据解析为 uint16
        endian:
            '>' 大端（Modbus 常见）
            '<' 小端
        """
        fmt = f"!{len(data_lst) // 2}H"
        ret = struct.unpack(fmt, bytes(data_lst))
        if struct.calcsize(fmt) == 2:
            ret = ret[0]
        return ret

    @staticmethod
    def change_word_order(data_lst):
        ret = []
        for i in range(0, len(data_lst), 4):
            ret.extend(data_lst[i + 2:i + 4])
            ret.extend(data_lst[i:i + 2])
        return ret

    def parse_bytes_to_f32_by_unpack_patch(self, data_lst):
        """解析字节为字符串"""
        change_word_order = self.change_word_order(data_lst)
        fmt = f"!{len(change_word_order) // 4}f"
        ret = struct.unpack(fmt, bytes(change_word_order))
        if struct.calcsize(fmt) == 4:
            ret = ret[0]
        return ret

    @staticmethod
    def parse_bytes_to_f32_by_pack(data):
        """解析字节为字符串"""
        fmt = "!f"
        pack_data_bytes = struct.pack(fmt, data)
        pack_data = pack_data_bytes[2:4] + pack_data_bytes[0:2]
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_f64_by_pack(data):
        """解析字节为字符串"""
        fmt = "!d"
        pack_data = struct.pack(fmt, data)
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_u16_by_pack(data):
        """解析字节为u16,2个字节"""
        fmt = "!H"
        pack_data = struct.pack(fmt, data)
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_u32_by_pack(data):
        """解析字节为u32,4个字节"""
        fmt = "!I"
        pack_data = struct.pack(fmt, data)
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_u16_by_pack_patch(*data):
        """解析字节为u16"""
        fmt = f"!{len(data)}H"
        pack_data = struct.pack(fmt, *data)
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_u32_by_pack_patch(*data):
        """解析字节为u32"""
        fmt = f"!{len(data)}I"
        pack_data = struct.pack(fmt, *data)
        return list(pack_data)

    def parse_bytes_to_f32_by_pack_patch(self, *data):
        """解析字节为f32"""
        pack_data = []
        _pack_data = None
        for _data in data:
            _pack_data = self.parse_bytes_to_f32_by_pack(_data)
            pack_data.extend(_pack_data)
        return list(pack_data)

    @staticmethod
    def parse_bytes_to_f64_by_pack_patch(*data):
        """解析字节为f64"""
        fmt = f"!{len(data)}d"
        pack_data = struct.pack(fmt, *data)
        return list(pack_data)

    def paser_harmonic_para_by_count(self, count, harmonic_data):
        """
        设置谐波参数帧打包
        count  : 谐波个数（0 = 清空）
        harmonic_data:
            (harmonic, amplitude, angle), ...
        """
        payload_len = 0
        payload = bytearray()
        _harmonic_data = [harmonic_data[i:i + 3] for i in range(0, len(harmonic_data), 3)]
        if count:
            if count != len(_harmonic_data):
                raise ValueError("count 与 harmonic_data 数量不一致")
            for item in _harmonic_data:
                harmonic, amplitude, angle = item
                if not (2 <= harmonic <= 100):
                    raise ValueError("Harmonic 次数范围 2~100")
                payload += bytes(self.parse_bytes_to_u16_by_pack(harmonic))
                payload += bytes(self.parse_bytes_to_f32_by_pack(amplitude))
                angle = math.radians(angle) if angle > 0 else math.radians(angle + 360)
                payload += bytes(self.parse_bytes_to_f32_by_pack(angle))
                # payload += struct.pack(
                #     '>Hff',  # uint8 + float + float
                #     harmonic,
                #     amplitude,
                #     angle
                # )
                payload_len += 5
        return list(payload), payload_len

    def paser_inter_harmonic_para_by_count(self, count, inter_harmonic):
        """
        设置3相谐波参数帧打包
        count  : 谐波个数（0 = 清空）
        harmonic_data:
            (harmonic, amplitude), ...
        """
        payload_len = 0
        payload = bytearray()
        _inter_harmonic_data = [inter_harmonic[i:i + 2] for i in range(0, len(inter_harmonic), 2)]
        if count:
            if count != len(_inter_harmonic_data):
                raise ValueError("count 与 harmonic_data 数量不一致")
            for item in _inter_harmonic_data:
                harmonic, amplitude = item
                harmonic = int(harmonic * 100)
                if not (0.25 * 100 <= harmonic <= 80.75 * 100):
                    raise ValueError("Harmonic 次数范围 0.25*100~80.75*100")
                if not (0 <= amplitude <= 0.4):
                    raise ValueError("amplitude 幅度范围 0~0.4")
                payload += bytes(self.parse_bytes_to_u16_by_pack(harmonic))
                payload += bytes(self.parse_bytes_to_f32_by_pack(amplitude))
                # payload += struct.pack(
                #     '>Hf',  # uint8 + float
                #     harmonic,
                #     amplitude,
                # )
                payload_len += 3
        return list(payload), payload_len

    def paser_slicker_para(self, slicker_para):
        """
        设置闪变参数帧打包
        slicker_para:
            (type, duty, modulation_depth, cpm), ...
        """
        if len(slicker_para) != 4 * 6:
            raise ValueError(f"slicker_para 输入参数个数不对,期望:4*6，实际:{len(slicker_para)}")
        payload_len = 0
        payload = bytearray()
        _slicker_para = [slicker_para[i:i + 4] for i in range(0, len(slicker_para), 4)]
        for item in _slicker_para:
            slicker_type, duty, modulation_depth, cpm = item
            if not (0 <= modulation_depth <= 0.4):
                raise ValueError("modulation_depth 次数范围 0~0.4")
            if not (0.00001 <= cpm <= 100):
                raise ValueError("cpm 幅度范围 0.00001Hz~100Hz")
            payload += bytes(self.parse_bytes_to_u16_by_pack(slicker_type))
            payload += bytes(self.parse_bytes_to_f32_by_pack_patch(duty, modulation_depth, cpm))
            # payload += struct.pack(
            #     '>H3f',  # uint16 + float + float + float
            #     slicker_type,
            #     duty,
            #     modulation_depth,
            #     cpm
            # )
            payload_len += 7
        return list(payload), payload_len

    def paser_swell_or_swag_para(self, swell_or_swag_para):
        """
        设置骤升/骤降参数帧打包
        swell_or_swag_para:
            (delay, ramp_in, period, ramp_out,change_to,end), ...
        """
        if len(swell_or_swag_para) != 6 * 6:
            raise ValueError(f"swell_or_swag_para 输入参数个数不对,期望:6*6，实际:{len(swell_or_swag_para)}")
        payload_len = 0
        payload = bytearray()
        _slicker_para = [swell_or_swag_para[i:i + 6] for i in range(0, len(swell_or_swag_para), 6)]
        for item in _slicker_para:
            delay, ramp_in, period, ramp_out, change_to, end = item
            if not (0.001 <= delay <= 60):
                raise ValueError("触发输入延时时间（单位 s,0.001s ~ 60s")
            if not (0.001 <= ramp_in <= 30):
                raise ValueError("斜入时间（单位 s，范围：0.001s ~ 30s）")
            if not (0.001 <= period <= 60):
                raise ValueError("骤升/骤降持续时间（单位 s，0.001s ~ 60s）")
            if not (0.001 <= ramp_out <= 30):
                raise ValueError("斜出时间（单位 s，0.001s ~ 30s）")
            if not (-0.4 <= change_to <= 0.4):
                raise ValueError("骤升/骤降幅值，起始电平的百分比（0 - ±0.4）")
            if not (0.001 <= end <= 60):
                raise ValueError("触发输出延时时间（单位 s, 0.001s ~ 60s")
            payload += bytes(self.parse_bytes_to_f32_by_pack_patch(delay, ramp_in, period, ramp_out, change_to, end))

            # payload += struct.pack(
            #     '>6f',  # float + float + float + float + float + float
            #     delay,
            #     ramp_in,
            #     period,
            #     ramp_out,
            #     change_to,
            #     end
            # )
            payload_len += 12
        return list(payload), payload_len

    # <editor-fold desc="三、系统信号">
    def get_model_ver(self):
        """
        联机命令，读取型号和版本号
        Model：终端产品型号，字符串表示，’\0’结束。如 Model=”6103”
        Ver.A，Ver.B：版本号，如 V1.00 表示为 Ver.A=01，Ver.B=00。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_ONLINE_LEN,
            CMD_ONLINE
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            model = self.parse_bytes_to_ascii_by_unpack(rcv_data[6:6 + 2])
            ver_a = self.parse_bytes_to_u32_by_unpack(rcv_data[8:8 + 4])
            ver_b = self.parse_bytes_to_u32_by_unpack(rcv_data[10:10 + 4])
            info = f"product model: {model}, ver.A: {ver_a}, ver.B: {ver_b}"
            return self.logger.info(info)

    # </editor-fold>

    # <editor-fold desc="四、标准源基本设置功能">
    def set_sour_freq(self, set_freq):
        """
        设置源输出频率命令（返回 OK）
        30~70Hz
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_FREQ_LEN,
            CMD_SET_FREQ
        )
        pack_data = self.parse_bytes_to_f32_by_pack(set_freq)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_wire_mode(self, set_wire=0):
        """
        设置源接线模式命令（返回 OK）
        WIRE：00 - 4 线（Y）,	01 - 3 线（V）
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_WIRE_LEN,
            CMD_SET_WIRE
        )
        pack_data = self.parse_bytes_to_u16_by_pack(set_wire)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name}, 指令正确接收 Success"
                self.logger.info(info)
                return

    def set_sour_gear(self, ua, ia, ub, ib, uc, ic):
        """
        设置源档位参数（返回 OK）
        其中 Ua、Ia、Ub、Ib、Uc、Ic 均为 1 个字的档位字。
        注意：当指定通道的设置档位与系统当前的档位不同时，系统会关闭该通道的源输出。
        注意设置时，Ua=Ub=Uc=U；Ia=Ib=Ic=I   如果各通道档位不一致，以 Ua 作为电压档位，Ia 作为电流档位。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_GEAR_LEN,
            CMD_SET_GEAR
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(ua, ia, ub, ib, uc, ic)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_sour_amp(self, ua, ia, ub, ib, uc, ic):
        """
        设置源幅度命令（返回 OK）
        其中 Ua、Ia、Ub、Ib、Uc、Ic 均为 2 个字的幅度值，浮点型数据。
        注意：发送源幅度的命令时，指定通道的幅值大于等于 0 时，系统会自动打开该通道的 源输出；幅值为负时（<0），则系统关闭该通道的源输出。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_AMP_LEN,
            CMD_SET_AMP
        )
        pack_data = self.parse_bytes_to_f32_by_pack_patch(ua, ia, ub, ib, uc, ic)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_sour_phase(self, ua_phase, ia_phase, ub_phase, ib_phase, uc_phase, ic_phase):
        """
        设置源相位命令（返回 OK）
        其中 Ua、Ia、Ub、Ib、Uc、Ic 均为 2 个字的相位值，浮点型数据。
        注意：Ua=0；当发送源相位的命令时，系统只改变输出波形，不会自动打开源输出。 相位的范围：±180°
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_PHASE_LEN,
            CMD_SET_PHASE
        )
        ua_rad = math.radians(ua_phase) if ua_phase > 0 else math.radians(ua_phase + 360)
        ia_rad = math.radians(ia_phase) if ia_phase > 0 else math.radians(ia_phase + 360)
        ub_rad = math.radians(ub_phase) if ub_phase > 0 else math.radians(ub_phase + 360)
        ib_rad = math.radians(ib_phase) if ib_phase > 0 else math.radians(ib_phase + 360)
        uc_rad = math.radians(uc_phase) if uc_phase > 0 else math.radians(uc_phase + 360)
        ic_rad = math.radians(ic_phase) if ic_phase > 0 else math.radians(ic_phase + 360)
        pack_data = self.parse_bytes_to_f32_by_pack_patch(ua_rad, ia_rad, ub_rad, ib_rad, uc_rad, ic_rad)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_active_power(self, channel, p_power):
        """
        设置有功功率命令（返回 OK）
        Channel：0-Pa，1-Pb，2-Pc，3-ΣP
        P：有功功率，为 2 个字的浮点型数据。当接线模式为 V 型时，Pb 自动屏蔽。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_P_POWER_LEN,
            CMD_SET_P_POWER
        )
        pack_data = self.parse_bytes_to_u16_by_pack(channel)
        pack_data.extend(
            self.parse_bytes_to_f32_by_pack(p_power)
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_reactive_power(self, channel, q_power):
        """
        设置无功功率命令（返回 OK）
        Channel：0-Qa，1-Qb，2-Qc，3-ΣQ
        Q：无功功率，为 2 个字的浮点型数据。当接线模式为 V 型时，Qb 自动屏蔽。
        注意：设置 Pa，Pb，Pc，Qa，Qb，Qc 时须保证当前通道的电压值不为 0；
        设置ΣP，ΣQ 时须保证三相电压中至少有一项不为 0。系统闭环包括对 P 和 Q 的闭环控制，
        在设置 P 和 Q 时须注意 Pa,Pb,Pc 可以同时设置，但是与ΣP 必须选一，Q 同 P。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_Q_POWER_LEN,
            CMD_SET_Q_POWER
        )
        pack_data = self.parse_bytes_to_u16_by_pack(channel)
        pack_data.extend(
            self.parse_bytes_to_f32_by_pack(q_power)
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_harmonic_para(self, channel: str, count, *harmonic_data):
        """
        设置谐波参数（返回 OK）
        Channel：通道选择，由位选择，D0-Ua,D1-Ia,D2-Ub,D3-Ib,D4-Uc,D5-Ic,	＝1 时有效
        Count：要设置的谐波个数(≥1)，协议长度 Length 随个数的变化而变化。
        Count =0 时清空指定通道的所有次数的谐波。
        Harmonic - 谐波次数，范围： 2 ~ 100
        Amplitude - 谐波幅度，范围： 0 ~ 0.4，2 个字的浮点型数据
        Angle - 谐波相位，范围： 0 ~ 2*PI (359.99)，2 个字的浮点型数据
        注意：该命令只设置谐波参数，并不输出谐波，只有当 0x007D 使能谐波输出。
        """
        frame_id = self.get_frame_id()
        ch_data = self.parse_bytes_to_u16_by_pack(int(channel, 2))
        cnt_data = self.parse_bytes_to_u16_by_pack(count)
        _harmonic_data = 0
        harmonic_data_len = 0
        if len(harmonic_data):
            _harmonic_data, harmonic_data_len = self.paser_harmonic_para_by_count(count, harmonic_data)
            pack_data = ch_data + cnt_data + _harmonic_data
        else:
            pack_data = ch_data + cnt_data
        # frame_id:1,length:1,cmd:1,checksum的长度为1,共计 +4
        cmd_set_harmonic_para_len = len([int(channel, 2)]) + len([count]) + harmonic_data_len + 4
        send_cmd = self.get_parse_data(
            cmd_set_harmonic_para_len,
            CMD_SET_HARMONIC_PARA
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_enable_harmonic_para(self, channel: str):
        """
        根据设置的谐波参数输出谐波（返回 OK）
        Channel：通道选择，由位选择，D0-Ua,D1-Ia,D2-Ub,D3-Ib,D4-Uc,D5-Ic,	＝1 时有效
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_ENABLE_HARMONIC_PARA_LEN,
            CMD_ENABLE_HARMONIC_PARA
        )
        pack_data = self.parse_bytes_to_u16_by_pack(int(channel, 2))
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_enable_sys_set_para(self, close_loop=0, close_type=0, power_type=0):
        """
        系统控制命令（返回 OK）
        CloseLoop 为闭环控制标志位（包括幅值闭环、相位闭环、功率闭环），1 个字的整型
        数据。0：允许闭环；1：禁止闭环。 注意：当任一通道转换到以下任何电能质量模式：闪变、间谐波、骤升/骤降，系统必 须禁止闭环；
        当所有通道恢复正常输出模式时，才能使能闭环控制。
        CloseType 为闭环控制类型设置，1 个字的整型数据。
        0：总有效值恒定模式，添加谐波时，通过降低基波分量使总有效值保持恒定；
        1：基波恒定模式，添加谐波时，基波恒定，总有效值变化。
        PowerType 为无功计算方法设置，1 个字的整型数据。
        0：滤波法 -> 时延法；1：三角法 -> 滤波法；2：时延法 -> 三角法。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_ENABLE_SYS_SET_PARA_LEN,
            CMD_ENABLE_SYS_SET_PARA
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(close_loop, close_type, power_type)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_enable_energy_verify(self, verify_type=0):
        """
        启动停止电能校验（返回 OK）
        Type – 电能校验类型，‘P’- 有功； ’Q’- 无功；0-停止
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_ENABLE_ENERGY_VERIFY_LEN,
            CMD_ENABLE_ENERGY_VERIFY
        )
        pack_data = self.parse_bytes_to_u16_by_pack(verify_type)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_energy_verify_para(self, spc, sqc, mpc, mqc, mdiv, mrou):
        """
        设置电能校验参数（自动停止电能输出和电能校验，返回 OK）
        SPC – 源有功电能常数，范围：1.0~599999999.0；
        SQC – 源无功电能常数，范围：1.0~599999999.0；
        MPC – 表有功电能常数，范围：1.0~599999999.0；
        MQC – 表无功电能常数，范围：1.0~599999999.0；
        以上四个参数保留 5 位小数，采用 2 个字的浮点数表示。
        MDIV – 分频系数，范围：1~9999999，2 个字的无符号整型数据，高字节在前；
        MROU – 校验圈数，范围：大于等于 1，2 个字的无符号整型数据，高字节在前。
        """
        spc = round(spc, 5)
        sqc = round(sqc, 5)
        mpc = round(mpc, 5)
        mqc = round(mqc, 5)
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_ENERGY_VERIFY_PARA_LEN,
            CMD_SET_ENERGY_VERIFY_PARA
        )
        pack_data = self.parse_bytes_to_f32_by_pack_patch(spc, sqc, mpc, mqc)
        pack_data.extend(
            self.parse_bytes_to_u32_by_pack_patch(mdiv, mrou)
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    # </editor-fold>

    # <editor-fold desc="五、电能质量标准源设置功能">
    def set_sys_mode(self, sys_mode=0):
        """
        设置系统模式（返回 OK）
        Mode：系统模式，0-正常输出模式；1-骤升/骤降模式；2-闪变模式；3-间谐波；
        10-校准模式。 注意：在某一时刻只能允许一种模式；但切换到相应的模式时，系统输出设置的参数。
        在 设置模式前须通过 5.2，5.3，5.4 命令设置相应通道模式的各项参数。发送命令前需进行合 法性检验。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_SYS_MODE_LEN,
            CMD_SET_SYS_MODE
        )
        pack_data = self.parse_bytes_to_u16_by_pack(sys_mode)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_flicker(self, *data):
        """
        设置闪变（返回 OK）
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_FLICKER_LEN,
            CMD_SET_FLICKER
        )
        pack_data, _ = self.paser_slicker_para(data)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_swell_or_swag(self, *data):
        """
        设置骤升/骤降（返回 OK）
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SET_SWELL_OR_SWAG_LEN,
            CMD_SET_SWELL_OR_SWAG
        )
        pack_data, _ = self.paser_swell_or_swag_para(data)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_inter_harmonic_para(self, channel: str, count, *harmonic_data):
        """
        设置间谐波参数（返回 OK）
        Channel：通道选择，由位选择，D0-Ua,D1-Ia,D2-Ub,D3-Ib,D4-Uc,D5-Ic,	＝1 时有效
        Count：要设置的间谐波个数(≥1)，协议长度 Length 随个数的变化而变化。
        Count =0 时清空指定通道的所有次数的间谐波。
        Harmonic – 间谐波次数，范围： 0.25 ~ 80.75  （乘以100，换成整数）
        Amplitude - 谐波幅度，范围： 0 ~ 0.4，2 个字的浮点型数据
        注意：该命令只设置间谐波参数，并不 输出间谐波，只有当 0x008D 使能间谐波输出。

        """
        frame_id = self.get_frame_id()
        ch_data = self.parse_bytes_to_u16_by_pack(int(channel, 2))
        cnt_data = self.parse_bytes_to_u16_by_pack(count)

        _harmonic_data = 0
        inter_harmonic_data_len = 0
        if len(harmonic_data):
            _inter_harmonic_data, inter_harmonic_data_len = self.paser_inter_harmonic_para_by_count(count,
                                                                                                    harmonic_data)
            pack_data = ch_data + cnt_data + _inter_harmonic_data
        else:
            pack_data = ch_data + cnt_data
        # frame_id:1,length:1,cmd:1,checksum的长度为1, 共计 +4
        cmd_set_inter_harmonic_para_len = len([int(channel, 2)]) + len([count]) + inter_harmonic_data_len + 4
        send_cmd = self.get_parse_data(
            cmd_set_inter_harmonic_para_len,
            CMD_SET_INTER_HARMONIC
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def set_enable_inter_harmonic_para(self, channel: str):
        """
        根据设置的间谐波参数输出间谐波（返回 OK）
        Channel：通道选择，由位选择，D0-Ua,D1-Ia,D2-Ub,D3-Ib,D4-Uc,D5-Ic,	＝1 时有效
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_ENABLE_INTER_HARMONIC_LEN,
            CMD_ENABLE_INTER_HARMONIC
        )
        pack_data = self.parse_bytes_to_u16_by_pack(int(channel, 2))
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    # </editor-fold>

    # <editor-fold desc="六、标准表功能">
    def get_meter_para(self):
        """
        读交流标准表参数
        WIRE 为 1 个字的整型数据；00 - 4 线（Y）,	01 - 3 线（V）
        CloseLoop 为 1 个字的整型数据：D3-D0：=0：允许闭环；=1：禁止闭环。
        D7-D4：=0：总有效值恒定；=1：基波恒定。
        D11-D8：=0：滤波法；=1：三角法；=2：时延法。
        EV 为 2 个字的浮点型数据，单位：%。
        ROUND 为当前校验圈数，2 个字的无符号整数数据，高字节在前。
        EV_Flag 为 1 个字的电能误差有效标志位，EV_Flag = 0 表示 EV 值无效；EV_Flag = ‘P’ 表示 EV 值为有功电能校验误差；
        EV_Flag = ‘Q’ 表示 EV 值为无功电能校验误差。读一次 当前电能误差后，有效标志位清零，当重新计算误差后，标志位有效。
        Code：备用代码。
        以上斜体表示的参数均为 2 个字的浮点型数据。
        （备注：CloseLoop实现有误：D11-D8：=0：时延法；=1：三角法；=1：滤波法）
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_GET_METER_PARA_LEN,
            CMD_GET_METER_PARA
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name}, 指令正确接收 Success")
                data_ret = rcv_data[6:-2]
                # fmt = "!2H1f6H28f1H2f1H14f"
                wire, close_loops = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[:struct.calcsize("!2H")]
                )
                close_loop = close_loops & 0x000F
                close_type = (close_loops & 0x00F0) >> 4
                power_type = (close_loops & 0x0F00) >> 8
                freq = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!2H"):struct.calcsize("!2H1f")]
                )
                ua_range, ia_range, ub_range, ib_range, uc_range, ic_range = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f"):struct.calcsize("!2H1f6H")]
                )
                (
                    ua, ia, ub, ib, uc, ic,
                    ua_angle, ia_angle, ub_angle, ib_angle, uc_angle, ic_angle,
                    pa, pb, pc, p_sys,
                    qa, qb, qc, q_sys,
                    sa, sb, sc, s_sys,
                    pf_a, pf_b, pf_c, pf_sys
                ) = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f6H"):struct.calcsize("!2H1f6H28f")]
                )
                ev_flag = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f6H28f"):struct.calcsize("!2H1f6H28f1H")]
                )
                ev, ev_round = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f6H28f1H"):struct.calcsize("!2H1f6H28f1H2f")]
                )
                code = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f6H28f1H2f"):struct.calcsize("!2H1f6H28f1H2f1H")]
                )
                (
                    fun_ua, fun_ia, fun_ub, fun_ib, fun_uc, fun_ic,
                    u_pos, u_neg, u_zero, u_unb,
                    i_pos, i_neg, i_zero, i_unb
                ) = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!2H1f6H28f1H2f1H"):]
                )

                ua_angle = math.degrees(ua_angle)
                ia_angle = math.degrees(ia_angle)
                ub_angle = math.degrees(ub_angle)
                ib_angle = math.degrees(ib_angle)
                uc_angle = math.degrees(uc_angle)
                ic_angle = math.degrees(ic_angle)
                self.logger.info(f"wire:{wire},"
                                 f"close_loop:{close_loop},"
                                 f"close_type:{close_type},"
                                 f"power_type:{power_type},"
                                 f"freq:{freq},"
                                 f"ua_range:{ua_range},"
                                 f"ia_range:{ia_range},"
                                 f"ub_range:{ub_range},"
                                 f"ib_range:{ib_range},"
                                 f"uc_range:{uc_range},"
                                 f"ic_range:{ic_range},"
                                 f"ua:{ua},"
                                 f"ia:{ia},"
                                 f"ub:{ub},"
                                 f"ib:{ib},"
                                 f"uc:{uc},"
                                 f"ic:{ic},"
                                 f"ua_angle:{ua_angle},"
                                 f"ia_angle:{ia_angle},"
                                 f"ub_angle:{ub_angle},"
                                 f"ib_angle:{ib_angle},"
                                 f"uc_angle:{uc_angle},"
                                 f"ic_angle:{ic_angle},"
                                 f"pa:{pa},"
                                 f"pb:{pb},"
                                 f"pc:{pc},"
                                 f"p_sys:{p_sys},"
                                 f"qa:{qa},"
                                 f"qb:{qb},"
                                 f"qc:{qc},"
                                 f"q_sys:{q_sys},"
                                 f"sa:{sa},"
                                 f"sb:{sb},"
                                 f"sc:{sc},"
                                 f"s_sys:{s_sys},"
                                 f"pf_a:{pf_a},"
                                 f"pf_b:{pf_b},"
                                 f"pf_c:{pf_c},"
                                 f"pf_sys:{pf_sys},"
                                 f"ev_flag:{ev_flag},"
                                 f"ev:{ev},"
                                 f"ev_round:{ev_round},"
                                 f"code:{code},"
                                 f"fun_ua:{fun_ua},"
                                 f"fun_ia:{fun_ia},"
                                 f"fun_ub:{fun_ub},"
                                 f"fun_ib:{fun_ib},"
                                 f"fun_uc:{fun_uc},"
                                 f"fun_ic:{fun_ic},"
                                 f"u_pos:{u_pos},"
                                 f"u_neg:{u_neg},"
                                 f"u_zero:{u_zero},"
                                 f"u_unb:{u_unb}"
                                 f"i_pos:{i_pos},"
                                 f"i_neg:{i_neg},"
                                 f"i_zero:{i_zero},"
                                 f"i_unb:{i_unb}"
                                 )
                return (
                    wire, close_loop, close_type, power_type, freq,
                    ua_range, ia_range, ub_range, ib_range, uc_range, ic_range,
                    ua, ia, ub, ib, uc, ic,
                    ua_angle, ia_angle, ub_angle, ib_angle, uc_angle, ic_angle,
                    pa, pb, pc, p_sys,
                    qa, qb, qc, q_sys,
                    sa, sb, sc, s_sys,
                    pf_a, pf_b, pf_c, pf_sys,
                    ev_flag, ev, ev_round, code,
                    fun_ua, fun_ia, fun_ub, fun_ib, fun_uc, fun_ic,
                    u_pos, u_neg, u_zero, u_unb,
                    i_pos, i_neg, i_zero, i_unb
                )

    def get_harmonic_para(self, channel, *harmonic_cnt):
        """
        读谐波参数
        N0 表示为第 N0 次谐波，……，Nm 表示为第 Nm 次谐波，均为 1 个字的整型数据。
        THDU 为电压总谐波畸变率，THDI 为电流总谐波畸变率，ΣP 为谐波总有功功率，ΣQ 为谐波总无功功率，
        HRUN0、PhUN0、HRIN0、PhIN0、P N0、Q N0 为第 N0 次谐波的电压含 有率，电压相位、电流含有率、电流相位、有功功率、无功功率。
        电压和电流的谐波相位 定义为谐波相对于基波的相位，均为 2 个字的浮点型数据。
        Channel：通道号，0-A 相，1-B 相，2-C 相。
        """
        frame_id = self.get_frame_id()
        ch_data = self.parse_bytes_to_u16_by_pack(channel)
        harmonic_cnt_data = self.parse_bytes_to_u16_by_pack_patch(*harmonic_cnt)
        pack_data = ch_data + harmonic_cnt_data
        # frame_id:1,length:1,cmd:1,checksum的长度为1, 共计 +4
        cmd_get_harmonic_para_len = len([channel]) + len([harmonic_cnt]) + 4
        send_cmd = self.get_parse_data(
            cmd_get_harmonic_para_len,
            CMD_GET_HARMONIC
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name},指令正确接收 Success")
                data_ret = rcv_data[6:-2]
                # fmt = f"!1H4f1H6f"
                ch = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[:struct.calcsize("!1H")]
                )
                thd_u, thd_i, p_sum, q_sum = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!1H"):struct.calcsize("!1H4f")]
                )
                num = self.parse_bytes_to_u16_by_unpack_patch(
                    data_ret[struct.calcsize("!1H4f"):struct.calcsize("!1H4f1H")]
                )
                hr_u, hr_u_angle, hr_i, hr_i_angle, hr_p, hr_q = self.parse_bytes_to_f32_by_unpack_patch(
                    data_ret[struct.calcsize("!1H4f1H"):]
                )
                hr_u_angle = math.degrees(hr_u_angle)
                hr_i_angle = math.degrees(hr_i_angle)
                self.logger.info(f"ch:{ch},"
                                 f"thd_u:{thd_u},"
                                 f"thd_i:{thd_i},"
                                 f"p_sum:{p_sum},"
                                 f"q_sum:{q_sum},"
                                 f"num:{num},"
                                 f"hr_u:{hr_u},"
                                 f"hr_u_angle:{hr_u_angle},"
                                 f"hr_i:{hr_i},"
                                 f"hr_i_angle:{hr_i_angle},"
                                 f"hr_p:{hr_p},"
                                 f"hr_q:{hr_q}"
                                 )
                return ch, thd_u, thd_i, p_sum, q_sum, num, hr_u, hr_u_angle, hr_i, hr_i_angle, hr_p, hr_q

    def get_3_phase_harmonic_para(self):
        """
        读三相谐波含有率参数
        Ua_N …Ic_N  该通道谐波的个数，有谐波值才上发，没有值谐波个数为0.
        THD 为谐波的总畸变率，
        xx_Har_cs 谐波次数，均为 1 个字的整型数据。
        xx_Har_hyl 谐波含有率，均为 2 个字的浮点型数据。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_GET_3_PHASE_HARMONIC_LEN,
            CMD_GET_3_PHASE_HARMONIC
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name},指令正确接收 Success")
                (
                    ua_n, ia_n, ub_n, ib_n, uc_n, ic_n,
                    ua_thd, ia_thd, ub_thd, ib_thd, uc_thd, ic_thd,
                ) = struct.unpack(f"!6H6f", bytes(rcv_data[6:6 * 12 + 6 * 4]))

                data_ret = rcv_data[6 * 12 + 6 * 4:-2]
                fmt_start = f"!"
                fmt_end = "1H2f" * sum([ua_n, ia_n, ub_n, ib_n, uc_n, ic_n])
                # ua_ic_cnt =
                # "1H2f" * ua_n + "1H2f" * ia_n +
                # "1H2f" * ub_n + "1H2f" * ib_n +
                # "1H2f" * uc_n + "1H2f" * ic_n
                fmt = fmt_start + fmt_end
                _data = struct.unpack(fmt, bytes(data_ret))
                for i in range(len(_data)):
                    return (
                        ua_n, ia_n, ub_n, ib_n, uc_n, ic_n,
                        ua_thd, ia_thd, ub_thd, ib_thd, uc_thd, ic_thd,
                        _data
                    )

    def get_3_phase_inter_harmonic_para(self):
        """
        读三相间谐波参数
        Ua_N …Ic_N  该通道谐波的个数，有谐波值才上发，没有值谐波个数为0.
        ETHD 为间谐波的总畸变率，
        Har_cs 间谐波次数，乘以100转换为整数，均为 1 个字的整型数据。
        Har_hyl 间谐波含有率，均为 2 个字的浮点型数据。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_GET_3_PHASE_INTER_HARMONIC_LEN,
            CMD_GET_3_PHASE_INTER_HARMONIC
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name},指令正确接收 Success")
                (
                    ua_n, ia_n, ub_n, ib_n, uc_n, ic_n,
                    ua_ethd, ia_ethd, ub_ethd, ib_ethd, uc_ethd, ic_ethd
                ) = struct.unpack(f"!6H6f", bytes(rcv_data[6:6 * 12 + 6 * 4]))
                data_ret = rcv_data[6 * 12 + 6 * 4:-2]
                fmt_start = f"!"
                fmt_end = "1H1f" * sum([ua_n, ia_n, ub_n, ib_n, uc_n, ic_n])
                fmt = fmt_start + fmt_end
                _data = struct.unpack(fmt, bytes(data_ret))
                return (
                    ua_n, ia_n, ub_n, ib_n, uc_n, ic_n,
                    ua_ethd, ia_ethd, ub_ethd, ib_ethd, uc_ethd, ic_ethd,
                    _data
                )

    def get_flicker_para(self):
        """
        读闪变参数
        其中：MUa,Mia,MUb,MIb,MUc,MIc 分别为当前各通道的输出模式；
        xx_Pst 为短时间闪变值，xx_Flag 为指标有效值标志位，=0：表示当前的 Pst 值是在严 格的矩形调制、
        230V/50Hz 和 120V/60Hz 条件下的试验数据值，即为真实 Pst 值；=1： Pst 为计算值，即“≈”。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_GET_FLICKER_LEN,
            CMD_GET_FLICKER
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name},指令正确接收 Success")
                data_ret = rcv_data[6:-2]
                # _data = struct.unpack("!7H6f6H", bytes(data_ret))
                _data = struct.unpack("!6H6f6H", bytes(data_ret))
                (
                    mua, mia, mub, mib, muc, mic,
                    ua_pst, ia_pst, ub_pst, ib_pst, uc_pst, ic_pst,
                    ua_flag, ia_flag, ub_flag, ib_flag, uc_flag, ic_flag,
                ) = _data
                return (
                    mua, mia, mub, mib, muc, mic,
                    ua_pst, ia_pst, ub_pst, ib_pst, uc_pst, ic_pst,
                    ua_flag, ia_flag, ub_flag, ib_flag, uc_flag, ic_flag,
                )

    def get_swell_or_swag_para(self):
        """
        读闪变参数
        其中：MUa,Mia,MUb,MIb,MUc,MIc 分别为当前各通道的输出模式，具体见表 4.1。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_GET_SWELL_OR_SWAG_LEN,
            CMD_GET_SWELL_OR_SWAG
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                self.logger.info(f"# {func_name},指令正确接收 Success")
                data_ret = rcv_data[6:-2]
                _data = struct.unpack("!6H36f", bytes(data_ret))
                (
                    mua, mia, mub, mib, muc, mic,
                    ua_delay, ua_ri, ua_p, ua_ro, ua_end, ua_am,
                    ia_delay, ia_ri, ia_p, ia_ro, ia_end, ia_am,
                    ub_delay, ub_ri, ub_p, ub_ro, ub_end, ub_am,
                    ib_delay, ib_ri, ib_p, ib_ro, ib_end, ib_am,
                    uc_delay, uc_ri, uc_p, uc_ro, uc_end, uc_am,
                    ic_delay, ic_ri, ic_p, ic_ro, ic_end, ic_am,
                ) = _data
                return (
                    mua, mia, mub, mib, muc, mic,
                    ua_delay, ua_ri, ua_p, ua_ro, ua_end, ua_am,
                    ia_delay, ia_ri, ia_p, ia_ro, ia_end, ia_am,
                    ub_delay, ub_ri, ub_p, ub_ro, ub_end, ub_am,
                    ib_delay, ib_ri, ib_p, ib_ro, ib_end, ib_am,
                    uc_delay, uc_ri, uc_p, uc_ro, uc_end, uc_am,
                    ic_delay, ic_ri, ic_p, ic_ro, ic_end, ic_am,
                )

    def send_error_code(self, error_code: str = "00111111"):
        """
        发送故障代码（系统一旦检测到故障代码，会自动发送，并进行保护操作）
        系统一旦检测到故障代码，会自动发送命令，并自动进行保护操作。
        保护操作如下：（1）关闭幅值输出，即幅值输出为 0；（2）关断档位输出，即所有的继 电器输出控制信号为 1。
        Error_Code：故障代码。D0-Ua,D1-Ia,D2-Ub,D3-Ib,D4-Uc,D5-Ic。=1 时，表示正常，=0 时表示该通道故障
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SEND_ERROR_CODE_LEN,
            CMD_SEND_ERROR_CODE
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(int(error_code, 2))
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    # </editor-fold>

    # <editor-fold desc="七、仪表校验命令格式">
    def switch_amp(self, urange, irange, level):
        """
        切换幅值校准点（切换档位和校准点；返回 OK）
        URange、IRange 为当前的档位，为 1 个字节的整型数， Level 为当前的校准点，0-零点，1-20%，2-100%，3-相位校准。
        建议 SUa=SIa=0；SUb=SIb=120°；SUc=SIc=240°。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_SWITCH_AMP_LEN,
            CMD_SWITCH_AMP
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(urange, irange, level)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def calibrate_amp(self, urange, irange, level, ua, ub, uc, ia, ib, ic):
        """
        幅值校准（校准输出值，返回 OK）
        URange、IRange 为当前的档位，为 1 个字节的整型数，
        注意：校准命令中的电压和电 流档位 URange，IRange 不执行换档动作，只是作为判断条件，判断校验的档位是否和当 前设置的档位一致。
        Level 为当前的校准点，0-零点，1-20%，2-100%，3-相位校准
        Ua、Ub、Uc、Ia、Ib、Ic 为当前所接的校准用标准表的读数，4 个字节的浮点 型数据。
        注意：输入的标准表读数限制在校准点标准值（SXx）的±20%之内，如果大于±20%，不发送校准命令。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_CALIBRATE_AMP_LEN,
            CMD_CALIBRATE_AMP
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(urange, irange, level)
        pack_data.extend(
            self.parse_bytes_to_f64_by_pack_patch(ua, ub, uc, ia, ib, ic)
        )
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def confirm_amp(self):
        """
        确认幅值校准（保存校准点数据；返回 OK）
        交流源校准后，退出校准模式前，发送该命令保持所有校准参数。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_CONFIRM_AMP_LEN,
            CMD_CONFIRM_AMP
        )
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    def reset_factory_setting(self, sour_type=0, urange=100, irange=100):
        """
        清空校准参数，恢复出厂默认设置（返回 OK）
        Type：类型，0-标准源， 1-标准表，2-钳形表，3-直流源，4-直流表（备用）。
        Urange、IRange：要清除的电压电流档位。注意，该命令会清除该档位下的幅值和相位 校准参数。
        当 Type=3 或 4 时，如果要清空电压档位参数，则设置电流档位为一个大于电流 总档位数的值（建议设置为 100）；
        如果要清空电流档位参数，则设置电压档位为一个大于 电压总档位数的值（建议设置为 100）。
        """
        frame_id = self.get_frame_id()
        send_cmd = self.get_parse_data(
            CMD_RESET_FACTORY_SETTING_LEN,
            CMD_RESET_FACTORY_SETTING
        )
        pack_data = self.parse_bytes_to_u16_by_pack_patch(sour_type, urange, irange)
        send_cmd.extend(pack_data)
        checksum = self.get_checksum(send_cmd)
        send_data = frame_id + send_cmd + checksum
        self.send(bytes(send_data))
        rcv_data = self.receive()
        if rcv_data:
            if list(rcv_data[4:4 + 2]) == send_cmd[2:2 + 2]:
                func_name = inspect.stack()[0].function
                info = f"# {func_name},指令正确接收 Success"
                return self.logger.info(info)

    # </editor-fold>


if __name__ == '__main__':
    so = DeviceUart(ser_com="COM18")

    # 三、系统信号
    # so.get_model_ver()
    # 七、仪表校验命令格式
    # so.switch_amp(urange=0, irange=0, level=0)
    # so.calibrate_amp(urange=0, irange=0, level=0, ua=1, ub=1, uc=1, ia=1, ib=1, ic=1)
    # so.confirm_amp()
    # so.reset_factory_setting(sour_type=0, urange=0, irange=0)

    # 四、标准源基本设置功能
    so.set_enable_sys_set_para(0, 0, 0)
    so.get_meter_para()

    # 重置系统参数
    so.set_sys_mode(sys_mode=0)
    # Y输出
    u = 0
    i = 1
    so.set_sour_gear(ua=u, ia=i, ub=u, ib=i, uc=u, ic=i)
    so.set_wire_mode(set_wire=0)
    so.set_sour_freq(set_freq=60)
    so.set_sour_phase(ua_phase=0, ia_phase=0, ub_phase=-120, ib_phase=-120, uc_phase=120, ic_phase=120)
    u_amp = 100
    i_amp = 2
    so.set_sour_amp(ua=u_amp, ia=i_amp, ub=u_amp, ib=i_amp, uc=u_amp, ic=i_amp)
    # Y切换V失败
    # so.set_wire_mode(set_wire=1) # Y -> V 切换失败
    # u_amp = 100 * (1 / math.sqrt(3))
    # i_amp = 2
    # so.set_sour_amp(u_amp, i_amp, u_amp, i_amp, u_amp, i_amp)
    # so.set_sys_mode(sys_mode=0)
    # so.set_enable_sys_set_para(0, 0, 2)
    # 设置有功、无功。屏显显示计算电流，Q:压流角不能为0°
    # so.set_active_power(channel=0, p_power=10.6)
    # so.set_reactive_power(channel=3, q_power=10.20)
    # 设置及使能谐波,谐波界面的屏幕会卡，屏显显示模糊
    # so.set_harmonic_para("00000001", 2, 2, 0.2, 0, 3, 0.2, 0)
    # so.set_enable_harmonic_para("00000001")
    # so.set_harmonic_para("00000010", 2, 2, 0.3, 0, 3, 0.3, 0)
    # so.set_enable_harmonic_para("00000010")
    # so.set_harmonic_para("00000100", 2, 2, 0.2, 0, 3, 0.2, 0)
    # so.set_enable_harmonic_para("00000100")
    # so.set_harmonic_para("00001000", 2, 2, 0.3, 0, 3, 0.3, 0)
    # so.set_enable_harmonic_para("00001000")
    # so.set_harmonic_para("00010000", 2, 2, 0.2, 0, 3, 0.2, 0)
    # so.set_enable_harmonic_para("00010000")
    # so.set_harmonic_para("00100000", 2, 2, 0.3, 0, 3, 0.3, 0)
    # so.set_enable_harmonic_para("00100000")
    # # 关闭谐波
    # so.set_harmonic_para("00111111", 0)
    # so.set_enable_harmonic_para("00111111")

    # 五、电能质量标准源设置功能
    # 设置使能闪变 -> 指令执行成功，设置参数解析错误，屏显解析错误
    # so.set_sys_mode(sys_mode=2)
    # so.set_flicker(
    #     0, 50, 0.1, 50,
    #     0, 50, 0.2, 50,
    #     0, 50, 0.3, 50,
    #     0, 50, 0.4, 50,
    #     0, 50, 0.5, 50,
    #     0, 50, 0.4, 50,
    # )
    # 设置使能骤变 -> 指令执行成功，设置参数解析错误，屏显解析错误
    # so.set_sys_mode(sys_mode=1)
    # so.set_swell_or_swag(
    #     1, 1, 2, 1, 0.1, 1,
    #     1, 1, 2, 1, 0.4, 1,
    #     1, 1, 2, 1, 0.2, 1,
    #     1, 1, 2, 1, 0.3, 1,
    #     1, 1, 2, 1, 0.3, 1,
    #     1, 1, 2, 1, 0.2, 1,
    # )
    # # 设置使能间谐波 -> pass。允许闭环且间谐波分量总和不超过基波幅值的40%，观察屏显3相幅值，幅值存在波动
    # so.set_sys_mode(sys_mode=3)
    # so.set_inter_harmonic_para(
    #     "00000001",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00000001")
    # so.set_inter_harmonic_para(
    #     "00000010",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00000010")
    # so.set_inter_harmonic_para(
    #     "00000100",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00000100")
    # so.set_inter_harmonic_para(
    #     "00001000",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00001000")
    # so.set_inter_harmonic_para(
    #     "00010000",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00010000")
    # so.set_inter_harmonic_para(
    #     "00100000",
    #     3,
    #     0.25, 0.1,
    #     0.50, 0.1,
    #     0.75, 0.1
    # )
    # so.set_enable_inter_harmonic_para("00100000")
    # # 关闭间谐波
    # so.set_sys_mode(sys_mode=3)
    # so.set_inter_harmonic_para("00111111", 0)
    # so.set_enable_inter_harmonic_para("00111111")

    # 六、标准表功能
    # 读取交流表参数 -> pass, 参数解析成功
    # so.get_meter_para()
    # 读取谐波参数 -> pass, 参数解析成功
    # so.get_harmonic_para(0, 2)
    # so.get_harmonic_para(1, 3)
    # so.get_harmonic_para(2, 2)
    # 读取3相谐波参数 -> pass, 参数解析失败, 实际返回参数个数与协议标注字段个数不同，咱无法使用该指令，可用读取谐波参数指令代替
    # so.get_3_phase_harmonic_para()
    # 读取3相间谐波参数  -> 指令返回为应答指令, 读取3相间谐波参数指令功能未做。
    # so.get_3_phase_inter_harmonic_para()
    # 读取闪变参数  -> pass, 参数解包数量对不上, 设置骤升/骤降参数指令的参数设置错误
    # so.get_flicker_para()
    # 读取骤升/骤降参数  -> pass, 设置骤升/骤降参数指令的参数设置错误
    # so.get_swell_or_swag_para()
    # 电压短接、电流开路会自动上报
    # so.send_error_code("00111111")

    so.close()
