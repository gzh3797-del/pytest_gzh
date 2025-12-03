import socket
import serial
import time
import logging
import pandas as pd
from pathlib import Path
from enum import Enum
from typing import Union, List, Dict
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

    def __init__(self, protocol: ModbusProtocol = ModbusProtocol.TCP):
        """
        初始化Modbus客户端

        Args:
            protocol: 协议类型，TCP或RTU，默认为TCP
        """
        self.protocol = protocol
        self.logger = logging.getLogger('ModbusClient')
        self._is_connected = False
        # 加载配置文件
        self._load_config()
        self.last_request_time = 0  # 记录上次请求时间
        self.min_interval = 0.2  # 最小间隔200ms
        if protocol == ModbusProtocol.TCP:
            self._init_tcp()
        elif protocol == ModbusProtocol.RTU:
            self._init_rtu()
        else:
            raise ValueError(f"不支持的协议类型: {protocol}")

        self.logger.info(f"Modbus客户端初始化完成 - 协议: {protocol.value}")

    def _load_config(self):
        """加载设备配置文件"""

        self.config = modbus_config

    def _init_tcp(self):
        """初始化TCP连接参数"""
        tcp_config = self.config["QT_tcp"]
        self.host = tcp_config["host"]
        self.port = tcp_config["port"]
        self.timeout = tcp_config["timeout"]
        self.slave_id = tcp_config["slave_id"]
        self.socket = None

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

        self.logger.info(f"RTU配置 - 串口: {self.port}, 波特率: {self.baudrate}, 从站ID: {self.slave_id}")

    def _ensure_connected(self):
        """确保连接已建立"""
        if self.socket is not None:
            # 检查socket是否真的有效
            try:
                # 简单测试socket是否可用
                self.socket.settimeout(0.1)  # 设置短暂超时
                self.socket.getpeername()  # 检查对端地址
                self.socket.settimeout(self.timeout)  # 恢复超时设置
                return  # 连接有效，直接返回
            except (OSError, socket.error, AttributeError):
                # 连接无效，关闭并重新连接
                self.logger.debug("现有连接无效，关闭并重新连接")
                try:
                    self.socket.close()
                except:
                    pass
                self.socket = None

        # 重新建立连接
        if self.socket is None:
            self.logger.info("正在建立新连接...")
            self.connect()

            # 验证新连接
            if self.socket is None:
                raise ConnectionError("无法建立连接: socket为None")

    def connect(self):
        """
        建立连接

        根据协议类型建立TCP连接或RTU串口连接
        """
        if self._is_connected:
            return

        if self.protocol == ModbusProtocol.TCP:
            self._connect_tcp()
        else:
            self._connect_rtu()

        self._is_connected = True

    def _connect_tcp(self):
        """建立TCP连接"""
        if self.socket is None:
            try:
                self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.socket.settimeout(self.timeout)
                self.socket.connect((self.host, self.port))
                self.logger.info(f"TCP连接成功: {self.host}:{self.port}")
            except socket.error as e:
                self.logger.error(f"TCP连接失败: {e}")
                raise

    def _connect_rtu(self):
        """建立RTU串口连接"""
        if self.serial_conn is None:
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
                self.logger.info(f"RTU串口连接成功: {self.port}, {self.baudrate}bps")
            except serial.SerialException as e:
                self.logger.error(f"RTU串口连接失败: {e}")
                raise

    def disconnect(self):
        """断开连接"""
        if not self._is_connected:
            return

        if self.protocol == ModbusProtocol.TCP:
            if self.socket:
                self.socket.close()
                self.socket = None
                self.logger.info("TCP连接已关闭")
        else:
            if self.serial_conn:
                self.serial_conn.close()
                self.serial_conn = None
                self.logger.info("RTU串口连接已关闭")

        self._is_connected = False

    def __enter__(self):
        """上下文管理器入口"""
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.disconnect()

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

    def _send_and_receive(self, request_data: bytes, max_retries: int = 10) -> str:
        """
        发送请求并接收响应，支持自动重连重试

        Args:
            request_data: 请求数据
            max_retries: 最大重试次数

        Returns:
            str: 格式化的16进制字符串，如 "00 01 00 00 00 07 01 03 04 00 00 00 3A"
        """
        last_exception = None

        for attempt in range(max_retries):
            try:
                self._ensure_connected()  # 确保连接已建立

                formatted_send = self._format_hex_string(request_data)
                self.logger.info(f"发送报文(尝试 {attempt + 1}/{max_retries}): {formatted_send}")

                # 发送和接收
                if self.protocol == ModbusProtocol.TCP:
                    self.socket.send(request_data)
                    response_bytes = self.socket.recv(1024)  # 接收的是字节

                    # ✅ 关键修改：返回格式化的16进制字符串，而不是解码的字符串
                    formatted_response = self._format_hex_string(response_bytes)
                    self.logger.info(f"接收报文: {formatted_response}")

                    # ✅ 返回格式化的16进制字符串
                    return formatted_response

            except (ConnectionResetError, ConnectionAbortedError,
                    socket.timeout, OSError) as e:
                last_exception = e
                self.logger.warning(f"连接异常，尝试重新连接 ({attempt + 1}/{max_retries}): {str(e)}")

                # 关闭现有连接
                if self.socket:
                    try:
                        self.socket.close()
                    except:
                        pass
                    self.socket = None

                # 等待后重试
                if attempt < max_retries:
                    time.sleep(1)  # 等待1秒后重试
                    continue

        # 所有重试都失败
        error_msg = f"发送请求失败，重试{max_retries}次后仍无法连接"
        self.logger.error(error_msg)
        raise ConnectionError(f"{error_msg}: {last_exception}")

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

    def parse_data(self, hex_string: str) -> int:
        """
        解析协议数据，返回数值结果

        协议格式说明（Modbus TCP）：
        事务标识（2字节）+ 协议标识（2字节）+ 长度（2字节）+ 设备地址（1字节）+ 功能码（1字节）+ 数据字节数（1字节）+ 数据（N字节）

        Args:
            hex_string: 十六进制字符串

        Returns:
            int: 解析出的数值
        """
        try:
            # 清理空格并转换为字节
            hex_clean = hex_string.replace(" ", "")
            data_bytes = bytes.fromhex(hex_clean)

            # 最小长度检查：事务标识2 + 协议标识2 + 长度2 + 设备地址1 + 功能码1 + 数据字节数1 = 9字节
            if len(data_bytes) < 9:
                return 0

            # 验证功能码是03（读取保持寄存器）
            if data_bytes[7] != 0x03:  # 第8个字节是功能码
                return 0

            # 获取数据字节数（第9个字节）
            data_length = data_bytes[8]

            # 检查是否有足够的数据
            if len(data_bytes) < 9 + data_length:
                return 0

            # 提取数据部分（从第10个字节开始）
            data_start_index = 9
            data_section = data_bytes[data_start_index:data_start_index + data_length]

            # 根据数据长度决定如何解析
            if data_length == 2:
                # 16位数据（2个字节）
                result = int.from_bytes(data_section, byteorder='big')
                return result
            elif data_length == 4:
                # 32位数据（4个字节）
                reg_high = int.from_bytes(data_section[0:2], byteorder='big')
                reg_low = int.from_bytes(data_section[2:4], byteorder='big')
                result = (reg_high << 16) | reg_low
                return result
            else:
                # 其他长度的数据，返回0
                return 0

        except Exception as e:
            print(f"解析错误: {e}")
            return 0


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
    with ModbusClient(ModbusProtocol.TCP) as client:
        client.validate_register_value('Cable  resistance', 0)
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
