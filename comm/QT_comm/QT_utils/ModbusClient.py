import socket
import serial
import time
import logging
import pandas as pd
from pathlib import Path
from enum import Enum
from typing import Union, List, Dict, Any, Optional
from modbus_config import modbus_config


class ModbusProtocol(Enum):
    """Modbus协议类型"""
    TCP = "TCP"
    RTU = "RTU"


class ModbusClient:
    """
    Modbus客户端类，支持TCP和RTU协议

    主要功能：
    - 支持Modbus TCP和RTU协议
    - 自动从Excel配置文件中查找寄存器地址和长度
    - 支持标准功能码(03/06/10)和自定义功能码(如6A)
    - 字符串自动转换为ASCII十六进制，长度不足时自动补零
    - 自动连接管理，用户无需手动调用connect/disconnect
    - 连接管理和错误处理
    """

    def __init__(self, protocol: ModbusProtocol = ModbusProtocol.TCP, **kwargs):
        """
        初始化Modbus客户端

        Args:
            protocol: 协议类型，TCP或RTU，默认为TCP
            **kwargs: 可选的配置参数，将覆盖默认配置
                    对于TCP: host, port, timeout, slave_id
                    对于RTU: port, baudrate, bytesize, parity, stopbits, timeout, slave_id
        """
        self.protocol = protocol
        self.logger = logging.getLogger('ModbusClient')
        self._is_connected = False
        # 加载配置文件
        self._load_config()
        self.last_request_time = 0  # 记录上次请求时间
        self.min_interval = 0.2  # 最小间隔200ms

        # 应用用户自定义配置（在初始化之前）
        self._override_config(kwargs)

        if protocol == ModbusProtocol.TCP:
            self._init_tcp()
        elif protocol == ModbusProtocol.RTU:
            self._init_rtu()
        else:
            raise ValueError(f"不支持的协议类型: {protocol}")

        self.logger.info(f"Modbus客户端初始化完成 - 协议: {protocol.value}")

    def _load_config(self):
        """加载设备配置文件"""
        self.config = modbus_config.copy()  # 创建副本，避免修改原始配置

    def _override_config(self, user_config: dict):
        """
        允许用户覆盖配置参数

        Args:
            user_config: 用户提供的配置参数
        """
        if not user_config:
            return

        if self.protocol == ModbusProtocol.TCP:
            config_key = "QT_tcp"
        else:
            config_key = "QT_rtu"

        # 记录用户提供的配置参数
        if user_config:
            self.logger.debug(f"用户提供的配置参数: {user_config}")

        # 更新配置字典
        for key, value in user_config.items():
            if key in self.config[config_key]:
                old_value = self.config[config_key][key]
                self.config[config_key][key] = value
                self.logger.info(f"配置覆盖: {key}: {old_value} -> {value}")
            else:
                self.logger.warning(f"忽略未知的配置参数: {key}")

    def _init_tcp(self):
        """初始化TCP连接参数"""
        tcp_config = self.config["QT_tcp"]
        self.host = tcp_config["host"]
        self.port = tcp_config["port"]
        self.timeout = tcp_config["timeout"]
        self.slave_id = tcp_config["slave_id"]
        self.socket = None
        self._max_connect_retries = 10  # 最大连接重试次数
        self._connect_retry_delay = 1.0  # 重试延迟（秒）

        self.logger.info(
            f"TCP配置 - 主机: {self.host}, 端口: {self.port}, 超时: {self.timeout}s, 从站ID: {self.slave_id}")

    def _init_rtu(self):
        """初始化RTU串口参数"""
        rtu_config = self.config["QT_rtu"]
        self.port = rtu_config["port"]
        self.baudrate = rtu_config["baudrate"]
        self.bytesize = rtu_config["bytesize"]
        self.parity = rtu_config["parity"]
        self.stopbits = rtu_config["stopbits"]
        self.timeout = rtu_config["timeout"]
        self.slave_id = rtu_config["slave_id"]
        self.serial_conn = None
        self._max_connect_retries = 5  # 串口连接重试次数较少
        self._connect_retry_delay = 2.0  # 串口重试延迟更长

        self.logger.info(
            f"RTU配置 - 串口: {self.port}, 波特率: {self.baudrate}, "
            f"数据位: {self.bytesize}, 校验位: {self.parity}, "
            f"停止位: {self.stopbits}, 超时: {self.timeout}s, 从站ID: {self.slave_id}")

    def _cleanup_socket(self):
        """清理socket连接并重置状态"""
        if self.socket:
            try:
                self.socket.close()
            except:
                pass
            self.socket = None
        self._is_connected = False

    def _cleanup_serial(self):
        """清理串口连接并重置状态"""
        if self.serial_conn:
            try:
                self.serial_conn.close()
            except:
                pass
            self.serial_conn = None
        self._is_connected = False

    def _ensure_connected(self, max_retries: int = None):
        """
        确保连接已建立，带有智能重试机制

        Args:
            max_retries: 最大重试次数，为None时使用默认值
        """
        if max_retries is None:
            max_retries = self._max_connect_retries

        retry_count = 0

        while retry_count < max_retries:
            try:
                if self.protocol == ModbusProtocol.TCP:
                    self._ensure_tcp_connected()
                else:
                    self._ensure_rtu_connected()

                # 连接成功
                return

            except (ConnectionError, socket.error, OSError, serial.SerialException) as e:
                retry_count += 1

                if retry_count >= max_retries:
                    self.logger.error(f"连接失败，已达最大重试次数 {max_retries}: {e}")
                    raise ConnectionError(f"无法建立连接，已达最大重试次数 {max_retries}: {e}")

                # 固定等待2秒（替代原来的指数退避）
                wait_time = 2.0  # 固定2秒等待时间
                self.logger.warning(f"连接失败 ({retry_count}/{max_retries})，等待 {wait_time:.1f} 秒后重试: {e}")

                # 清理连接
                if self.protocol == ModbusProtocol.TCP:
                    self._cleanup_socket()
                else:
                    self._cleanup_serial()

                time.sleep(wait_time)

    def _ensure_tcp_connected(self):
        """确保TCP连接已建立"""
        # 检查现有连接是否有效
        if self.socket is not None and self._is_connected:
            try:
                # 快速测试socket是否可用
                self.socket.settimeout(0.5)

                # 方法1：尝试获取对端地址
                try:
                    self.socket.getpeername()
                    self.socket.settimeout(self.timeout)
                    self.logger.debug("现有TCP连接有效")
                    return
                except (OSError, socket.error, AttributeError):
                    pass

                # 方法2：尝试发送空数据
                try:
                    # 设置非阻塞模式测试
                    self.socket.setblocking(False)
                    try:
                        self.socket.send(b'')
                        self.logger.debug("TCP连接测试通过(send)")
                    except BlockingIOError:
                        # 非阻塞模式下缓冲区满也是正常情况
                        self.logger.debug("TCP连接测试通过(缓冲区满)")
                    except (OSError, socket.error):
                        raise ConnectionError("TCP连接已断开")
                    finally:
                        self.socket.setblocking(True)
                        self.socket.settimeout(self.timeout)
                    return

                except (OSError, socket.error):
                    raise ConnectionError("TCP连接测试失败")

            except Exception as e:
                self.logger.warning(f"TCP连接测试失败: {e}")
                # 继续下面的重新连接流程

        # 需要建立新连接
        self.logger.info("建立新的TCP连接...")

        # 清理旧连接
        self._cleanup_socket()

        # 创建新连接
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.settimeout(self.timeout)

            # 连接
            self.socket.connect((self.host, self.port))

            # 设置keepalive
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)

            self._is_connected = True
            self.logger.info(f"TCP连接成功: {self.host}:{self.port}")

        except socket.error as e:
            self._cleanup_socket()
            raise ConnectionError(f"TCP连接失败: {e}")

    def _ensure_rtu_connected(self):
        """确保RTU串口连接已建立"""
        # 检查现有连接是否有效
        if self.serial_conn is not None and self._is_connected:
            try:
                # 检查串口是否打开
                if self.serial_conn.is_open:
                    # 尝试读取串口状态
                    self.serial_conn.in_waiting
                    self.logger.debug("现有RTU串口连接有效")
                    return
            except (serial.SerialException, AttributeError, OSError) as e:
                self.logger.warning(f"RTU串口连接测试失败: {e}")
                # 继续下面的重新连接流程

        # 需要建立新连接
        self.logger.info("建立新的RTU串口连接...")

        # 清理旧连接
        self._cleanup_serial()

        # 创建新连接
        try:
            parity_map = {'N': serial.PARITY_NONE, 'E': serial.PARITY_EVEN, 'O': serial.PARITY_ODD}
            stopbits_map = {1: serial.STOPBITS_ONE, 2: serial.STOPBITS_TWO}

            self.serial_conn = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                bytesize=self.bytesize,
                parity=parity_map.get(self.parity, serial.PARITY_NONE),
                stopbits=stopbits_map.get(self.stopbits, serial.STOPBITS_ONE),
                timeout=self.timeout
            )

            # 等待串口稳定
            time.sleep(0.1)

            # 清空缓冲区
            self.serial_conn.reset_input_buffer()
            self.serial_conn.reset_output_buffer()

            self._is_connected = True
            self.logger.info(f"RTU串口连接成功: {self.port}, {self.baudrate}bps, parity={self.parity}")

        except serial.SerialException as e:
            self._cleanup_serial()
            raise ConnectionError(f"RTU串口连接失败: {e}")

    def connect(self):
        """
        建立连接（兼容性方法）
        根据协议类型建立TCP连接或RTU串口连接
        """
        self._ensure_connected(max_retries=10)# 只尝试一次，保持原有行为

    def _connect_tcp(self):
        """建立TCP连接（兼容性方法）"""
        self._ensure_tcp_connected()

    def _connect_rtu(self):
        """建立RTU串口连接（兼容性方法）"""
        self._ensure_rtu_connected()

    def disconnect(self):
        """断开连接"""
        if self.protocol == ModbusProtocol.TCP:
            self._cleanup_socket()
            self.logger.info("TCP连接已关闭")
        else:
            self._cleanup_serial()
            self.logger.info("RTU串口连接已关闭")

    def __enter__(self):
        """上下文管理器入口"""
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.disconnect()

    # 以下方法保持不变...
    def _find_register_info(self, description: str) -> tuple:
        """
        根据描述查找寄存器地址和长度
        """
        try:
            config_dir = Path(__file__).parent.parent / "config"
            if not config_dir.exists():
                raise ValueError(f"config目录不存在: {config_dir}")

            xlsx_files = list(config_dir.glob("*.xlsx"))
            if not xlsx_files:
                raise ValueError(f"config目录下未找到.xlsx文件: {config_dir}")

            excel_file = xlsx_files[0]
            self.logger.info(f"查找寄存器信息: {description}, 文件: {excel_file}")

            excel_file_obj = pd.ExcelFile(excel_file)
            sheet_names = excel_file_obj.sheet_names

            for i in range(2, len(sheet_names)):
                sheet_name = sheet_names[i]
                try:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if df.empty:
                        continue
                    try:
                        mask = df['Descrption'] == description
                    except:
                        mask = df['Description'] == description
                    matching_rows = df[mask]

                    if len(matching_rows) > 0:
                        row = matching_rows.iloc[0]
                        address_hex = row['Start(Hex)']
                        reg_length = row['Reg']

                        # 格式化地址
                        if isinstance(address_hex, int):
                            address_hex = f"{address_hex:04X}"
                        elif isinstance(address_hex, str):
                            if address_hex.startswith('0x') or address_hex.startswith('0X'):
                                address_hex = address_hex[2:].upper()
                            address_hex = address_hex.zfill(4).upper()

                        # 确保reg_length是整数
                        if not isinstance(reg_length, int):
                            try:
                                reg_length = int(reg_length)
                            except (ValueError, TypeError):
                                reg_length = 1

                        self.logger.info(
                            f"找到寄存器信息: {description} -> 地址:{address_hex}, 寄存器长度:{reg_length} (工作表: {sheet_name})")
                        excel_file_obj.close()
                        return address_hex, reg_length

                except Exception as e:
                    self.logger.warning(f"读取工作表 '{sheet_name}' 失败: {e}")
                    continue

            excel_file_obj.close()
            raise ValueError(f"在所有工作表中未找到描述为 '{description}' 的寄存器")

        except Exception as e:
            self.logger.error(f"查找寄存器信息失败: {e}")
            raise

    def _calculate_crc16(self, data: bytes) -> bytes:
        """
        计算RTU CRC16校验码
        """
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc = (crc >> 1) ^ 0xA001
                else:
                    crc = crc >> 1
        return crc.to_bytes(2, 'little')

    def _format_hex_string(self, data: Union[bytes, str]) -> str:
        """
        格式化十六进制字符串为可读格式（带空格）
        """
        if isinstance(data, str):
            data = data.replace(' ', '')
            try:
                data = bytes.fromhex(data)
            except ValueError:
                return data

        hex_str = data.hex().upper() if isinstance(data, bytes) else data
        return ' '.join(hex_str[i:i + 2] for i in range(0, len(hex_str), 2))

    def _value_to_hex(self, value: Union[int, str, None], reg_length: int) -> str:
        """
        将值转换为十六进制字符串

        Args:
            value: 要转换的值，支持整数、字符串或None
            reg_length: 寄存器长度(从Excel的Reg列获取)

        Returns:
            str: 十六进制字符串
        """
        if value is None:
            return ""

        if isinstance(value, int):
            # 整数处理：支持大整数拆分
            if value > 0xFFFF:
                # 对于超过16位的整数，拆分为高16位和低16位
                id_high = (value >> 16) & 0xFFFF
                id_low = value & 0xFFFF
                hex_str = f"{id_high:04X}{id_low:04X}"

                # 如果还需要更多寄存器，补零
                if reg_length > 2:
                    hex_str += '0000' * (reg_length - 2)
            else:
                # 对于16位及以下的整数
                if reg_length == 1:
                    hex_str = f"{value:04X}"
                else:
                    # 第一个寄存器为0，第二个寄存器为值，其余补零
                    hex_str = f"0000{value:04X}"
                    if reg_length > 2:
                        hex_str += '0000' * (reg_length - 2)

            return hex_str

        elif isinstance(value, str):
            # 字符串处理：短字符串向前补零，长字符串向后补零

            # 先转换为十六进制
            hex_str = ""
            for char in value:
                hex_str += f"{ord(char):02X}"

            # 计算需要的总字符数（每个寄存器4个十六进制字符）
            required_chars = reg_length * 4
            current_chars = len(hex_str)

            self.logger.debug(f"字符串转换: '{value}' -> {hex_str}, 需要{required_chars}字符, 当前{current_chars}字符")

            # 判断是短字符串还是长字符串
            if current_chars <= 4:
                # 短字符串：向前补零（右对齐）
                if current_chars < required_chars:
                    padding_chars = required_chars - current_chars
                    hex_str = '0' * padding_chars + hex_str  # 向前补零
                    self.logger.debug(f"短字符串向前补零: 添加{padding_chars}个零 -> {hex_str}")
            else:
                # 长字符串：向后补零（左对齐）
                if current_chars < required_chars:
                    padding_chars = required_chars - current_chars
                    hex_str += '0' * padding_chars  # 向后补零
                    self.logger.debug(f"长字符串向后补零: 添加{padding_chars}个零 -> {hex_str}")
                # 长度超过时截断
                elif current_chars > required_chars:
                    hex_str = hex_str[:required_chars]
                    self.logger.debug(f"字符串截断: 保留{required_chars}字符 -> {hex_str}")

            return hex_str

        else:
            raise ValueError(f"不支持的值类型: {type(value)}，只支持int、str或None")

    def _build_modbus_request(self, function_code: str, address_hex: str,
                              reg_length: int, value: Union[int, str, None] = None) -> bytes:
        """
        构建Modbus请求报文
        """
        # 构建PDU（协议数据单元）
        if value is None:
            # 读取请求
            pdu_hex = f"{self.slave_id:02X}{function_code}{address_hex}{reg_length:04X}"
        else:
            # 写入请求
            value_hex = self._value_to_hex(value, reg_length)

            if function_code == '06':
                # 写单个寄存器
                pdu_hex = f"{self.slave_id:02X}{function_code}{address_hex}{value_hex}"
            elif function_code == '6A':
                # 自定义功能码6A
                byte_count = len(value_hex) // 2  # 计算字节数
                pdu_hex = f"{self.slave_id:02X}{function_code}{address_hex}{reg_length:04X}{byte_count:02X}{value_hex}"
            else:
                # 写多个寄存器（功能码10）
                byte_count = len(value_hex) // 2
                pdu_hex = f"{self.slave_id:02X}{function_code}{address_hex}{reg_length:04X}{byte_count:02X}{value_hex}"

        # 根据协议类型构建完整报文
        if self.protocol == ModbusProtocol.TCP:
            transaction_id = 0x0001
            protocol_id = 0x0000
            pdu_bytes = bytes.fromhex(pdu_hex[2:])
            length_field = len(pdu_bytes) + 1
            mbap_header = f"{transaction_id:04X}{protocol_id:04X}{length_field:04X}{self.slave_id:02X}"
            full_message_hex = mbap_header + pdu_hex[2:]
        else:
            full_message_hex = pdu_hex

        # 转换为bytes
        request_data = bytes.fromhex(full_message_hex)

        # RTU需要添加CRC
        if self.protocol == ModbusProtocol.RTU:
            crc = self._calculate_crc16(request_data)
            request_data += crc

        self.logger.debug(f"构建请求: 功能码={function_code}, 地址={address_hex}, 寄存器长度={reg_length}, 值={value}")

        return request_data

    def _send_and_receive(self, request_data: bytes) -> str:
        """
        发送请求并接收响应

        Args:
            request_data: 请求数据

        Returns:
            str: 格式化的16进制字符串，如 "00 01 00 00 00 07 01 03 04 00 00 00 3A"
        """
        try:
            # 确保连接已建立（包含重试机制）
            self._ensure_connected()

            formatted_send = self._format_hex_string(request_data)
            self.logger.info(f"发送报文: {formatted_send}")

            # 发送和接收
            if self.protocol == ModbusProtocol.TCP:
                self.socket.send(request_data)
                response_bytes = self.socket.recv(1024)

                formatted_response = self._format_hex_string(response_bytes)
                self.logger.info(f"接收报文: {formatted_response}")

                return formatted_response
            else:
                # RTU模式
                self.serial_conn.write(request_data)
                # 根据请求长度计算预期响应长度
                time.sleep(0.05)  # 等待设备响应
                response_bytes = self.serial_conn.read(1024)

                formatted_response = self._format_hex_string(response_bytes)
                self.logger.info(f"接收报文: {formatted_response}")

                return formatted_response

        except (ConnectionResetError, ConnectionAbortedError,
                socket.timeout, OSError, serial.SerialException) as e:
            self.logger.error(f"发送/接收失败: {e}")

            # 标记连接为断开状态
            if self.protocol == ModbusProtocol.TCP:
                self._cleanup_socket()
            else:
                self._cleanup_serial()

            raise ConnectionError(f"通信失败: {e}")

    def _ensure_interval(self):
        """确保请求间隔不大于200ms（限制最大间隔）"""
        current_time = time.time()
        elapsed = current_time - self.last_request_time
        if elapsed > self.min_interval:
            self.logger.warning(f"请求间隔 {elapsed:.3f}s 超过最大限制 {self.min_interval}s")
        # 不需要sleep，只是记录警告
        self.last_request_time = time.time()

    def validate_register_value(self, description: str, value: Union[int, str, None] = None,
                                use_6a: bool = False) -> str:
        """
        寄存器值校验（支持读取和写入操作）

        根据参数自动选择功能码：
        - value=None, use_6a=False: 读取操作，功能码03
        - value!=None, use_6a=False: 写入操作，功能码10
        - value!=None, use_6a=True: 自定义写入，功能码6A

        Args:
            description: 寄存器描述，用于查找地址和寄存器长度(Reg)
            value: 要写入的值，支持整数、字符串，None表示读取操作
            use_6a: 是否使用自定义功能码6A，仅对写入操作有效

        Returns:
            str: 接收到的报文
        """

        # 查找寄存器地址和寄存器长度(Reg)
        address_hex, reg_length = self._find_register_info(description)

        # 根据参数自动选择功能码
        if value is None:
            function_code = '03'  # 读取
            self.logger.info(f"执行读取操作: {description}, 地址: {address_hex}, 寄存器长度: {reg_length}")
        else:
            if use_6a:
                function_code = '6A'  # 自定义功能码
                self.logger.info(
                    f"执行6A写入操作: {description}, 地址: {address_hex}, 寄存器长度: {reg_length}, 值: {value}")
            else:
                function_code = '10'  # 写多个寄存器
                self.logger.info(
                    f"执行写入操作: {description}, 地址: {address_hex}, 寄存器长度: {reg_length}, 值: {value}")

        # 构建请求报文
        request_data = self._build_modbus_request(function_code, address_hex, reg_length, value)

        # 发送并接收响应
        return self._send_and_receive(request_data)

    # 以下方法保持不变...
    def send_custom_message(self, message_hex: Union[str, List[str]],
                            delay_between_messages: float = 0.2) -> Union[str, List[Dict]]:
        """
        发送自定义报文（支持TCP和RTU协议，支持单个或批量发送）

        Args:
            message_hex: 十六进制报文字符串（可包含空格）或字符串列表
                        在RTU模式下会自动添加CRC校验码
            delay_between_messages: 批量发送时的消息间隔（秒）

        Returns:
            str 或 list: 单个报文返回字符串，批量报文返回结果列表
        """
        if isinstance(message_hex, str):
            return self._send_single_custom_message(message_hex)
        elif isinstance(message_hex, list):
            return self._send_batch_custom_messages(message_hex, delay_between_messages)
        else:
            raise ValueError("message_hex 必须是字符串或字符串列表")

    def _send_single_custom_message(self, message_hex: str) -> str:
        """
        发送单个自定义报文

        在RTU模式下会自动计算并添加CRC校验码
        """
        # 清理输入，移除空格
        message_hex = message_hex.replace(' ', '')

        try:
            # 转换为bytes
            request_data = bytes.fromhex(message_hex)

            # 在RTU模式下，如果报文没有CRC，自动添加CRC
            if self.protocol == ModbusProtocol.RTU:
                # 检查是否已经有CRC（通常RTU报文最后2字节是CRC）
                if len(request_data) >= 2:
                    # 计算当前数据的CRC
                    calculated_crc = self._calculate_crc16(request_data[:-2])
                    current_crc = request_data[-2:]

                    # 如果CRC不匹配，重新计算整个报文的CRC
                    if calculated_crc != current_crc:
                        # 重新计算整个报文的CRC（假设输入的是没有CRC的报文）
                        calculated_crc = self._calculate_crc16(request_data)
                        request_data = request_data + calculated_crc
                        self.logger.debug("自动添加CRC校验码到自定义报文")
                else:
                    # 报文太短，直接计算CRC
                    calculated_crc = self._calculate_crc16(request_data)
                    request_data = request_data + calculated_crc
                    self.logger.debug("为短报文自动添加CRC校验码")

            # 发送并接收响应
            return self._send_and_receive(request_data)

        except ValueError as e:
            self.logger.error(f"自定义报文格式错误: {e}")
            raise

    def _send_batch_custom_messages(self, message_list: List[str],
                                    delay_between_messages: float = 0.2) -> List[Dict]:
        """
        批量发送自定义报文
        """
        results = []

        for i, message_hex in enumerate(message_list):
            try:
                self.logger.info(f"发送批量报文 [{i + 1}/{len(message_list)}]")
                response = self._send_single_custom_message(message_hex)
                results.append({
                    'index': i,
                    'message': message_hex,
                    'response': response,
                    'status': 'success'
                })

                if i < len(message_list) - 1:
                    time.sleep(delay_between_messages)

            except Exception as e:
                self.logger.error(f"批量报文 [{i + 1}] 发送失败: {e}")
                results.append({
                    'index': i,
                    'message': message_hex,
                    'response': None,
                    'status': 'error',
                    'error': str(e)
                })

                if i < len(message_list) - 1:
                    time.sleep(delay_between_messages)

        return results

    def batch_validate(self, commands: List[Dict],
                       delay_between_commands: float = 0.2) -> Dict[str, str]:
        """
        批量寄存器值校验

        Args:
            commands: 命令列表，每个命令为字典格式：
                     [{'description': 'Year', 'value': 2023, 'use_6a': False}, ...]
                     value为None表示读取操作
            delay_between_commands: 命令间隔（秒）

        Returns:
            dict: 结果字典 {description: 收到的报文}
        """
        results = {}

        for i, cmd in enumerate(commands):
            description = cmd.get('description')
            value = cmd.get('value')
            use_6a = cmd.get('use_6a', False)

            if not description:
                continue

            try:
                self.logger.info(f"执行批量校验 [{i + 1}/{len(commands)}]: {description}")
                response = self.validate_register_value(description, value, use_6a)
                results[description] = response

            except Exception as e:
                results[description] = None
                self.logger.error(f"❌ 寄存器 '{description}' 校验失败: {e}")

            if i < len(commands) - 1:
                time.sleep(delay_between_commands)

        return results

    def compare_results(self, results1: Dict, results2: Dict) -> bool:
        """
        对比两个结果字典是否一致

        Args:
            results1: 第一个结果字典 {description: 报文}
            results2: 第二个结果字典 {description: 报文}

        Returns:
            bool: 是否完全一致
        """
        # 检查是否有相同的key
        if set(results1.keys()) != set(results2.keys()):
            self.logger.error("❌ 结果字典的key不一致")
            print("❌ 结果字典的key不一致")
            return False

        all_same = True

        for description in results1.keys():
            value1 = results1[description]
            value2 = results2[description]

            # 处理None值
            if value1 is None and value2 is None:
                self.logger.info(f"✅ {description}: 两者都为None")
                print(f"✅ {description}: 两者都为None")
                continue
            elif value1 is None or value2 is None:
                self.logger.error(f"❌ {description}: 一个为None，另一个不为None")
                print(f"❌ {description}: 一个为None，另一个不为None")
                all_same = False
                continue

            # 去除空格后比较
            clean1 = value1.replace(' ', '')
            clean2 = value2.replace(' ', '')

            if clean1 == clean2:
                self.logger.info(f"✅ {description}: 一致")
                print(f"✅ {description}: 一致")
            else:
                self.logger.error(f"❌ {description}: 不一致")
                print(f"❌ {description}: 不一致")
                print(f"   第一次: {value1}")
                print(f"   第二次: {value2}")
                all_same = False

        if all_same:
            self.logger.info("✅ 所有结果完全一致")
            print("✅ 所有结果完全一致")
        else:
            self.logger.error("❌ 存在不一致的结果")
            print("❌ 存在不一致的结果")

        return all_same

    def keep_alive(self):
        """
        保持连接活跃（可选方法）
        如果需要长时间保持连接，可以定期调用此方法
        """
        if not self._is_connected:
            self.connect()

    def parse_data(self, hex_string: str, data_type: str = 'auto', protocol: str = None) -> Union[
        int, str, bytes, None]:
        """
        解析Modbus响应数据，支持TCP和RTU协议

        Args:
            hex_string: 十六进制字符串，如 "01 03 20 21 28 37 41 35 37 30 31 46 36 45 46 ..."
            data_type: 数据类型，可选值：
                      'auto' - 自动检测（默认）
                      'int' - 整数
                      'uint16' - 16位无符号整数
                      'uint32' - 32位无符号整数
                      'mac' - MAC地址（自动检测ASCII或二进制格式）
                      'mac_ascii' - ASCII十六进制格式的MAC地址
                      'mac_binary' - 二进制格式的MAC地址
                      'ascii' - ASCII字符串
                      'hex' - 原始十六进制字符串
                      'bytes' - 字节数组
                      'raw' - 原始数据（不解析协议头）
            protocol: 协议类型，'tcp' 或 'rtu'，为None时自动检测

        Returns:
            Union[int, str, bytes, None]: 解析结果
        """
        try:
            # 清理空格
            hex_clean = hex_string.replace(" ", "")

            if not hex_clean:
                return None

            data_bytes = bytes.fromhex(hex_clean)

            # 如果指定了协议类型，按指定协议解析
            if protocol:
                if protocol.lower() == 'tcp':
                    return self._parse_modbus_tcp(data_bytes, data_type)
                elif protocol.lower() == 'rtu':
                    return self._parse_modbus_rtu(data_bytes, data_type)
                else:
                    raise ValueError(f"不支持的协议类型: {protocol}")

            # 自动检测协议类型
            if self._is_modbus_tcp(data_bytes):
                return self._parse_modbus_tcp(data_bytes, data_type)
            elif self._is_modbus_rtu(data_bytes):
                return self._parse_modbus_rtu(data_bytes, data_type)
            else:
                # 不是标准Modbus格式，直接解析数据
                return self._parse_raw_data(data_bytes, data_type)

        except Exception as e:
            self.logger.error(f"解析错误: {e}")
            return None

    def _is_modbus_tcp(self, data_bytes: bytes) -> bool:
        """检查是否是Modbus TCP报文"""
        if len(data_bytes) < 8:
            return False

        # Modbus TCP头部：事务ID(2) + 协议ID(2) + 长度(2) + 单元ID(1) + 功能码(1)
        # 协议ID应为0
        protocol_id = int.from_bytes(data_bytes[2:4], byteorder='big')
        return protocol_id == 0

    def _is_modbus_rtu(self, data_bytes: bytes) -> bool:
        """检查是否是Modbus RTU报文"""
        if len(data_bytes) < 4:
            return False

        # RTU报文至少包含：设备地址(1) + 功能码(1) + 数据 + CRC(2)
        # 功能码应在有效范围内
        function_code = data_bytes[1]
        valid_function_codes = [0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x0F, 0x10]
        return function_code in valid_function_codes

    def _parse_modbus_tcp(self, data_bytes: bytes, data_type: str) -> Union[int, str, bytes, None]:
        """解析Modbus TCP报文"""
        if len(data_bytes) < 9:
            return None

        # 解析TCP头部
        transaction_id = int.from_bytes(data_bytes[0:2], byteorder='big')
        protocol_id = int.from_bytes(data_bytes[2:4], byteorder='big')
        length = int.from_bytes(data_bytes[4:6], byteorder='big')
        unit_id = data_bytes[6]  # 设备地址
        function_code = data_bytes[7]

        # 检查长度是否匹配
        if len(data_bytes) < 6 + length:
            return None

        # 数据部分从第8个字节开始
        if len(data_bytes) >= 9:
            data_length = data_bytes[8]
            if len(data_bytes) >= 9 + data_length:
                data_section = data_bytes[9:9 + data_length]
            else:
                data_section = data_bytes[9:]
        else:
            data_section = data_bytes[8:]

        self.logger.debug(f"TCP报文解析 - 事务ID:{transaction_id}, 长度:{length}, "
                          f"设备:{unit_id}, 功能码:{function_code:02X}, 数据长度:{len(data_section)}")

        return self._parse_data_section(data_section, data_type, function_code)

    def _parse_modbus_rtu(self, data_bytes: bytes, data_type: str) -> Union[int, str, bytes, None]:
        """解析Modbus RTU报文"""
        if len(data_bytes) < 4:
            return None

        # 解析RTU头部
        unit_id = data_bytes[0]  # 设备地址
        function_code = data_bytes[1]

        # 数据部分从第2个字节开始（不包括最后的CRC）
        if len(data_bytes) >= 3:
            # RTU没有明确的数据长度字段，需要根据功能码判断
            if function_code in [0x03, 0x04]:  # 读取保持/输入寄存器
                if len(data_bytes) >= 5:
                    # 字节数在第三个字节
                    data_length = data_bytes[2]
                    data_section = data_bytes[3:3 + data_length]
                else:
                    data_section = data_bytes[2:-2]  # 排除CRC
            else:
                data_section = data_bytes[2:-2]  # 排除CRC
        else:
            data_section = bytes()

        self.logger.debug(f"RTU报文解析 - 设备:{unit_id}, 功能码:{function_code:02X}, "
                          f"数据长度:{len(data_section)}")

        return self._parse_data_section(data_section, data_type, function_code)

    def _parse_raw_data(self, data_bytes: bytes, data_type: str) -> Union[int, str, bytes, None]:
        """解析原始数据（非标准Modbus格式）"""
        return self._parse_data_section(data_bytes, data_type, None)

    def _parse_data_section(self, data_section: bytes, data_type: str,
                            function_code: int = None) -> Union[int, str, bytes, None]:
        """解析数据部分"""
        if len(data_section) == 0:
            return None

        # 自动检测数据类型
        if data_type == 'auto':
            if function_code in [0x03, 0x04]:  # 读取寄存器
                if len(data_section) == 2:
                    data_type = 'uint16'
                elif len(data_section) == 4:
                    data_type = 'uint32'
                elif len(data_section) >= 12 and self._is_ascii_hex(data_section[:12]):
                    data_type = 'mac_ascii'
                else:
                    data_type = 'hex'
            else:
                # 检查是否是ASCII十六进制格式
                if len(data_section) >= 12 and self._is_ascii_hex(data_section[:12]):
                    data_type = 'mac_ascii'
                elif all(32 <= b <= 126 or b == 0 for b in data_section[:20]):
                    data_type = 'ascii'
                else:
                    data_type = 'hex'

        # 根据数据类型解析
        if data_type == 'int' or data_type == 'uint16':
            if len(data_section) >= 2:
                return int.from_bytes(data_section[:2], byteorder='big')

        elif data_type == 'uint32':
            if len(data_section) >= 4:
                return int.from_bytes(data_section[:4], byteorder='big')

        elif data_type == 'mac':
            # 自动检测MAC格式
            result = self._parse_mac_ascii(data_section)
            if result:
                return result
            return self._parse_mac_binary(data_section)

        elif data_type == 'mac_ascii':
            return self._parse_mac_ascii(data_section)

        elif data_type == 'mac_binary':
            return self._parse_mac_binary(data_section)

        elif data_type == 'ascii':
            return self._parse_ascii_string(data_section)

        elif data_type == 'hex':
            return data_section.hex().upper()

        elif data_type == 'bytes':
            return data_section

        elif data_type == 'raw':
            # 返回原始数据，不进行解析
            return data_section

        else:
            return data_section.hex().upper()

    def _is_ascii_hex(self, data_bytes: bytes) -> bool:
        """检查是否是ASCII十六进制格式（每个字节都是0-9,A-F的ASCII码）"""
        for byte in data_bytes:
            if not (48 <= byte <= 57 or 65 <= byte <= 70 or 97 <= byte <= 102 or byte == 0):
                return False
        return True

    def _parse_mac_ascii(self, data_section: bytes) -> str:
        """解析ASCII十六进制格式的MAC地址"""
        try:
            hex_chars = []
            for byte in data_section:
                if byte == 0:  # 遇到空字符停止
                    break
                if 48 <= byte <= 57:  # 0-9
                    hex_chars.append(chr(byte))
                elif 65 <= byte <= 70:  # A-F
                    hex_chars.append(chr(byte))
                elif 97 <= byte <= 102:  # a-f
                    hex_chars.append(chr(byte).upper())
                else:
                    # 非十六进制字符，停止收集
                    break

            hex_str = ''.join(hex_chars)

            # 如果是MAC地址，应该有12个十六进制字符
            if len(hex_str) >= 12:
                return hex_str[:12]
            elif len(hex_str) > 0:
                return hex_str
            else:
                return ""

        except Exception as e:
            self.logger.error(f"解析ASCII MAC错误: {e}")
            return ""

    def _parse_mac_binary(self, data_section: bytes, byte_order: str = 'big') -> str:
        """解析二进制格式的MAC地址"""
        try:
            if len(data_section) >= 6:
                mac_bytes = data_section[:6]
                if byte_order == 'little':
                    mac_bytes = bytes(reversed(mac_bytes))
                return ':'.join(f'{b:02X}' for b in mac_bytes)
            return ""
        except Exception as e:
            self.logger.error(f"解析二进制MAC错误: {e}")
            return ""

    def _parse_ascii_string(self, data_section: bytes) -> str:
        """解析ASCII字符串"""
        try:
            ascii_str = data_section.decode('ascii', errors='ignore')
            clean_str = ''.join(c for c in ascii_str if 32 <= ord(c) <= 126 or c == '\n' or c == '\r' or c == '\t')
            null_index = clean_str.find('\x00')
            if null_index >= 0:
                clean_str = clean_str[:null_index]
            return clean_str.rstrip('\x00')
        except Exception as e:
            self.logger.error(f"解析ASCII字符串错误: {e}")
            return ""

    # 专门解析MAC地址的方法（更简单易用）
    def parse_mac_address(self, hex_string: str, protocol: str = None) -> str:
        """
        专门解析MAC地址，支持TCP和RTU协议

        Args:
            hex_string: 十六进制字符串
            protocol: 协议类型，'tcp' 或 'rtu'，为None时自动检测

        Returns:
            str: MAC地址字符串（如D46D6DF4DE88）
        """
        result = self.parse_data(hex_string, 'mac', protocol)
        if result:
            # 如果返回的是带冒号的格式，去掉冒号
            if isinstance(result, str) and ':' in result:
                return result.replace(':', '')
            return str(result)
        return ""

    def get_config(self) -> dict:
        """
        获取当前配置

        Returns:
            dict: 当前配置信息
        """
        if self.protocol == ModbusProtocol.TCP:
            return self.config["QT_tcp"].copy()
        else:
            return self.config["QT_rtu"].copy()


# 使用示例
# 使用示例
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,  # 设置日志级别为DEBUG，可以看到所有日志
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),  # 输出到控制台
        ]
    )
    # with ModbusClient(ModbusProtocol.TCP) as rtu_client:
    #     r = rtu_client.validate_register_value('Serial Number', 'DE55061235', True)
    with ModbusClient(ModbusProtocol.RTU) as client:
        client.validate_register_value('Serial Number', 'DE55051235', True)
        # client.send_custom_message('01 10 10 02 00 01 02 00 03')
        # client.validate_register_value('Cable  resistance', 0)
        # for _ in range(20):
        #     client.send_custom_message('00 08 00 00 00 09 01 10 10 16 00 01 02 00 00')
        #     client.send_custom_message('00 08 00 00 00 06 01 03 10 16 00 01')
        #     client.send_custom_message('00 08 00 00 00 06 01 03 10 3F 00 01')

    #     coms = ['01 10 40 00 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 04 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 08 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 0C 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 10 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 14 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 18 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 1C 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 20 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 24 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 28 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 2C 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 30 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 34 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 38 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 3C 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 40 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 44 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 48 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 4C 00 04 08 40 00 00 00 00 00 00 00',
    #             '01 10 40 50 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 54 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 58 00 04 08 40 10 00 00 00 00 00 00',
    #             '01 10 40 5C 00 04 08 40 10 00 00 00 00 00 00',
    #             ]
    #     rtu_client.send_custom_message(coms)