import pandas as pd
import os
import glob
import sys

# 打包命令： pyinstaller --onefile --console --icon=tools/modbus报文生成/icon.ico  tools/modbus报文生成/modbus.py

# ===== 路径处理：保证打包后也能找到同目录的文件 =====
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))


class ModbusCommandGenerator:
    def __init__(self, slave_address=0x01):
        self.slave_addr = slave_address
        self.register_map = {}  # 存储参数名到(地址, 寄存器数量)的映射
        self.special_params = []

    def load_modbus_table(self, excel_path: str):
        """
        加载Modbus Address表 - 修复版：优先检查开关触发条件
        """
        try:
            xls = pd.ExcelFile(excel_path)
            sheet_names = xls.sheet_names
            registered_params = 0

            for sheet_name in sheet_names[2:]:
                print(f"处理工作表: {sheet_name}")
                try:
                    df = pd.read_excel(excel_path, sheet_name=sheet_name)
                except:
                    continue
                if df.empty:
                    continue

                # --- 1. 确定 Block 列 ---
                # 策略：优先找名字含 'block' 的列，找不到则默认使用第一列（索引0）
                block_col_name = None
                for col in df.columns:
                    if 'block' in str(col).lower():
                        block_col_name = col
                        break

                # 如果没找到叫 block 的列，且表格有列，就假设第一列是 Block 信息列 (针对 Unnamed: 0 情况)
                if block_col_name is None and len(df.columns) > 0:
                    block_col_name = df.columns[0]
                    # print(f"  未找到Block列名，默认使用第一列: {block_col_name}")

                # --- 2. 初始化开关 ---
                f6A_active = False

                # 列名映射检查 (保持原逻辑)
                found_columns = {'Start(Hex)': 'Start(Hex)', 'Description': 'Description', 'Reg': 'Reg'}

                # 这里的 Start(Hex) 如果列名不匹配，请根据实际 Excel 修改，比如有的表可能有空格
                # 简单的模糊匹配列名逻辑 (可选)
                for col in df.columns:
                    c_str = str(col).strip()
                    if 'Start' in c_str and 'Hex' in c_str: found_columns['Start(Hex)'] = col
                    if 'Description' in c_str: found_columns['Description'] = col
                    if 'Reg' in c_str: found_columns['Reg'] = col

                reg_column = found_columns['Reg'] if found_columns['Reg'] in df.columns else None

                # --- 3. 遍历行 ---
                for row_idx, row in df.iterrows():
                    try:
                        if not f6A_active:
                            # 获取 Block 列的值
                            block_val = row.get(block_col_name)
                            if block_val is not None:
                                val_str = str(block_val)
                                # 检查是否包含 product info
                                if 'product info' in val_str.lower():
                                    print(f"  >>> 在第 {row_idx + 1} 行找到 Product Info，开启提取模式")
                                    f6A_active = True

                        # 【常规数据解析】：如果没有地址数据，跳过（但开关状态已保留）
                        start_hex = row.get(found_columns['Start(Hex)'])

                        # 如果地址为空，说明这行可能是标题行或空行，直接跳过
                        if pd.isna(start_hex) or str(start_hex).strip() in ['', 'nan']:
                            continue

                        # 解析地址
                        if isinstance(start_hex, str):
                            addr_str = start_hex.strip().upper().replace("'", "").replace(" ", "").replace("0X", "")
                        else:
                            addr_str = str(int(start_hex))
                        try:
                            address = int(addr_str, 16)
                        except ValueError:
                            # 如果地址解析失败，跳过
                            continue

                        # 解析参数名
                        desc = row.get(found_columns['Description'])
                        if pd.isna(desc) or str(desc).strip() in ['', 'nan']:
                            continue
                        param_name = str(desc).strip()

                        # 解析寄存器数量
                        reg_count = 1
                        if reg_column:
                            reg_val = row.get(reg_column)
                            if not (pd.isna(reg_val) or str(reg_val).strip() in ['', 'nan']):
                                try:
                                    reg_count = int(reg_val)
                                except:
                                    pass

                        # 存入通用字典
                        self.register_map[param_name] = (address, reg_count)
                        registered_params += 1

                        if f6A_active:
                            self.special_params.append(param_name.strip().lower())

                    except Exception as e:
                        # print(f"  处理行 {row_idx + 1} 异常: {str(e)}")
                        continue

            print(f"已成功注册 {registered_params} 个参数")
        except Exception as e:
            print(f"加载Modbus表时发生错误: {str(e)}")
            raise

    def get_param_name(self, param_name: str):
        # 搜索参数名（不区分大小写模糊匹配）
        param_match = []
        cleaned_name = param_name.strip().lower()
        # 检查所有参数名（模糊匹配）
        for name in self.register_map:
            if cleaned_name == name.lower():
                param_match = [name]
                break
            if cleaned_name in name.lower():
                param_match.append(name)

        if not param_match:
            # 没有找到匹配项，显示错误
            available_params = list(self.register_map.keys())
            if not available_params:
                raise ValueError("注册表中没有可用参数")

            # 尝试查找最接近的匹配
            possible_matches = []
            for name in available_params:
                if cleaned_name in name.lower():
                    possible_matches.append(name)

            if possible_matches:
                error_msg = f"参数 '{param_name}' 未精确定义。可能匹配:\n"
                for match in possible_matches[:5]:
                    error_msg += f"- {match}\n"
                if len(possible_matches) > 5:
                    error_msg += f"... 共找到 {len(possible_matches)} 个可能匹配"
            else:
                error_msg = f"参数 '{param_name}' 未在配置表中找到。请检查名称。"

            raise ValueError(error_msg)
        return param_match

    def generate_command(self, param_match: str, param_value):
        """
        生成Modbus指令 - 自动识别数据类型 + 寄存器数量取Reg列
        """
        address, reg_count = self.register_map[param_match]

        # 根据参数名称选择功能码
        function_code = 0x6A if param_match.lower() in self.special_params else 0x10
        # 针对 mac / serial number 强制规则
        if param_match.lower() in ["serial number", "mac address(port1)", "mac address(port2)", "mac address(port3)", "mac address"]:
            reg_count = 6
            required_bytes = 12
        else:
            required_bytes = reg_count * 2

        # 处理参数值（自动识别数据类型）
        data_bytes = self._process_parameter(param_value)

        # 计算需要的数据字节数
        required_bytes = reg_count * 2

        # 检查数据长度是否足够
        if len(data_bytes) > required_bytes:
            raise ValueError(f"输入数据过长: 需要{required_bytes}字节, 实际{len(data_bytes)}字节")

        # 数据长度不足时，在前面补0
        if len(data_bytes) < required_bytes:
            # 只有非特殊参数才补零
            if function_code == 0x10:
                padding = [0] * (required_bytes - len(data_bytes))
                data_bytes = padding + data_bytes
                print(f"  数据长度不足，已补{len(padding)}个0")

        # 数据字节数 = 寄存器数量 × 2
        data_length = reg_count * 2

        # 构造写入命令帧
        command_frame = [
            self.slave_addr,
            function_code,
            *self._split_16bit(address),
            *self._split_16bit(reg_count),
            required_bytes,
            *data_bytes
        ]
        write_cmd = ' '.join(f"{byte:02X}" for byte in command_frame)

        # === 构造读取命令帧 ===
        read_frame = [
            self.slave_addr,
            0x03,  # 功能码固定读取保持寄存器
            *self._split_16bit(address),
            *self._split_16bit(reg_count)
        ]
        read_cmd = ' '.join(f"{byte:02X}" for byte in read_frame)

        # === 构造TCP写入命令帧 ===
        # TCP头部: 00 00 00 00 00 + RTU指令长度
        tcp_write_length = len(command_frame)
        tcp_write_header = [0x00, 0x00, 0x00, 0x00, 0x00, tcp_write_length]
        tcp_write_frame = tcp_write_header + command_frame
        tcp_write_cmd = ' '.join(f"{byte:02X}" for byte in tcp_write_frame)

        # === 构造TCP读取命令帧 ===
        tcp_read_length = len(read_frame)
        tcp_read_header = [0x00, 0x00, 0x00, 0x00, 0x00, tcp_read_length]
        tcp_read_frame = tcp_read_header + read_frame
        tcp_read_cmd = ' '.join(f"{byte:02X}" for byte in tcp_read_frame)

        return write_cmd, read_cmd, tcp_write_cmd, tcp_read_cmd, param_match

    def _split_16bit(self, value):
        """将16位值拆分为高低字节"""
        return [(value >> 8) & 0xFF, value & 0xFF]

    def _process_parameter(self, param_value):
        """
        参数值处理 - 返回数据字节列表（不进行补零）
        """
        # 自动判断类型
        if isinstance(param_value, int):
            # 数值类型处理
            value = param_value
            # 转换为大端序字节数组（2字节）
            return [(value >> 8) & 0xFF, value & 0xFF]
        elif isinstance(param_value, str):
            # 检查是否是数字字符串
            value_str = param_value.strip()
            # 尝试解析为整数
            try:
                if value_str.startswith('0x'):
                    # 十六进制格式
                    value = int(value_str[2:], 16)
                else:
                    # 十进制格式
                    value = int(value_str)
                # 数值类型处理
                return [(value >> 8) & 0xFF, value & 0xFF]
            except:
                # 字符串类型处理
                # 转换为ASCII字节
                ascii_bytes = param_value.encode('ascii')
                return list(ascii_bytes)
        else:
            return self._process_parameter(str(param_value))

    def list_parameters(self):
        """列出所有注册的参数"""
        params = list(self.register_map.keys())
        if not params:
            print("未注册任何参数")
            return

        print(f"已注册参数 ({len(params)} 个):")
        for i, param in enumerate(params):
            address, reg_count = self.register_map[param]
            print(f"{i + 1}. {param} (地址: {hex(address)}, 寄存器数: {reg_count})")


# 主程序
if __name__ == "__main__":
    print("=== Modbus指令生成器 (寄存器数量取Reg列) ===")
    generator = ModbusCommandGenerator()
    # 自动查找当前目录下的xlsx文件
    current_dir = os.getcwd()
    xlsx_files = glob.glob(os.path.join(current_dir, '*.xlsx'))

    if not xlsx_files:
        print("错误: 在当前目录下未找到任何xlsx文件!")
        sys.exit(1)
    # 使用找到的第一个xlsx文件
    excel_file = xlsx_files[0]
    print(f"自动选择Modbus表: {os.path.basename(excel_file)}")
    try:
        generator.load_modbus_table(excel_file)
        print("配置表加载成功")
        generator.list_parameters()
        print(f"功能码为6A的参数列表: {generator.special_params}")
    except Exception as e:
        print(f"配置表加载失败: {str(e)}")
        sys.exit(1)
    print("\n输入指令生成模式 (输入 'q' 退出)")

    while True:
        try:
            param_name = input("\n请输入参数名: ").strip()
            if not param_name:
                continue
            if param_name.lower() in ['q', 'quit', 'exit']:
                print("退出程序")
                break
            param_name = generator.get_param_name(param_name)
            if len(param_name) > 1:
                print(f"找到多个参数{param_name}, 请重新输入!")
                continue
            print(f"找到参数: {param_name}")
            param_value = input("请输入参数值: ").strip()
            if not param_value:
                print("参数值不能为空")
                continue
            # 生成指令（自动识别数据类型）
            cmd = generator.generate_command(param_name[0], param_value)
            print(f"RTU写入指令: {cmd[0]}")
            print(f"RTU读取指令: {cmd[1]}")
            print(f"TCP写入指令: {cmd[2]}")
            print(f"TCP读取指令: {cmd[3]}")

        except ValueError as e:
            print(f"错误: {str(e)}")
        except Exception as e:
            print(f"发生错误: {str(e)}")