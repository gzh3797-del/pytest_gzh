import pandas as pd
import re
import os
import glob
import sys

# 打包命令：pyinstaller --onefile --console --icon=icon.ico  modbus.py


# ===== 路径处理：保证打包后也能找到同目录的文件 =====
if getattr(sys, 'frozen', False):
    # 打包成 exe 后
    base_path = os.path.dirname(sys.executable)
else:
    # 源码运行
    base_path = os.path.dirname(os.path.abspath(__file__))


class ModbusCommandGenerator:
    def __init__(self, slave_address=0x01):
        self.slave_addr = slave_address
        self.register_map = {}  # 存储参数名到(地址, 寄存器数量)的映射

    def load_modbus_table(self, excel_path: str):
        """
        加载Modbus Address表 - 解析Reg列(寄存器数量)
        """
        try:
            # 读取所有Excel工作表
            xls = pd.ExcelFile(excel_path)
            sheet_names = xls.sheet_names
            # 遍历除了前两个sheet之外的所有sheet
            registered_params = 0

            for sheet_name in sheet_names[2:]:
                print(f"处理工作表: {sheet_name}")
                try:
                    df = pd.read_excel(excel_path, sheet_name=sheet_name)
                except:
                    print(f"  读取工作表失败")
                    continue
                if df.empty:
                    print(f"  工作表为空")
                    continue

                # 尝试检测列名
                found_columns = {}

                # 检查可能的列名组合 - 新增Reg列
                found_columns = {
                    'Start(Hex)': 'Start(Hex)',
                    'Description': 'Description',
                    'Reg': 'Reg'
                }

                # for col_key, possible_names in possible_columns.items():
                #     for name in possible_names:
                #         if name in df.columns:
                #             found_columns[col_key] = name
                #             break

                # 检查是否找到了必要的列
                if not found_columns.get('Start(Hex)') or not found_columns.get('Description'):
                    print(f"  缺少必要列，跳过此表。找到的列: {list(df.columns)}")
                    continue

                # 如果没有找到Reg列，则使用默认值1
                reg_column = found_columns.get('Reg')
                if not reg_column:
                    print("  警告: 未找到Reg列，默认寄存器数量为1")

                # 处理数据行
                for row_idx, row in df.iterrows():
                    try:
                        # 获取地址值
                        start_hex = row[found_columns['Start(Hex)']]
                        if pd.isna(start_hex) or str(start_hex).strip() in ['', 'nan']:
                            continue
                        # 标准化地址字符串
                        if isinstance(start_hex, str):
                            addr_str = start_hex.strip().upper()
                            # 移除可能的引号、空格和0x前缀
                            addr_str = addr_str.replace("'", "").replace(" ", "").replace("0X", "")
                        else:
                            addr_str = str(int(start_hex))

                        # 解析地址
                        try:
                            address = int(addr_str, 16)  # 优先尝试16进制解析
                        except:
                            address = int(addr_str)  # 尝试10进制解析

                        # 获取参数名称
                        desc = row[found_columns['Description']]
                        if pd.isna(desc) or str(desc).strip() in ['', 'nan']:
                            continue

                        param_name = str(desc).strip()

                        # 获取寄存器数量
                        if reg_column:
                            reg_count = row[reg_column]
                            if pd.isna(reg_count) or str(reg_count).strip() in ['', 'nan']:
                                reg_count = 1  # 默认值
                            else:
                                try:
                                    reg_count = int(reg_count)
                                except:
                                    reg_count = 1
                        else:
                            reg_count = 1

                        # 存储到字典: (地址, 寄存器数量)
                        self.register_map[param_name] = (address, reg_count)
                        # print(f"  注册参数: '{param_name}' (地址={hex(address)}, 寄存器数={reg_count})")
                        registered_params += 1
                    except Exception as e:
                        print(f"  处理行 {row_idx + 1} 失败: {str(e)}")
                        continue

            print(f"已成功注册 {registered_params} 个参数")
            if registered_params == 0:
                print("警告: 没有找到任何有效参数! 请检查Excel文件格式")

        except Exception as e:
            print(f"加载Modbus表时发生错误: {str(e)}")
            raise

    def generate_command(self, param_name: str, param_value):
        """
        生成Modbus指令 - 自动识别数据类型 + 寄存器数量取Reg列
        """
        # 搜索参数名（不区分大小写模糊匹配）
        param_match = None
        cleaned_name = param_name.strip().lower()

        # 检查所有参数名（模糊匹配）
        for name in self.register_map:
            if cleaned_name in name.lower():
                param_match = name
                break

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

        # 获取寄存器地址和寄存器数量
        address, reg_count = self.register_map[param_match]

        # 定义需要使用特殊功能码0x6A的参数列表
        special_params = {
            "hardware version",
            "function model type",
            "reserved 1",
            "mac address",
            "certification type",
            "aiao or dido flag",
            "reserved 2",
            "hardware patch number",
            "Serial Number"
        }
        # 根据参数名称选择功能码
        function_code = 0x6A if param_match.lower() in special_params else 0x10
        # ✅ 针对 mac / serial number 强制规则
        if param_match.lower() in ["serial number", "mac address"]:
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

        return write_cmd, read_cmd

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
    # 初始化生成器
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

        # 列出所有参数（可选）
        # list_params = input("是否列出所有参数? (y/n): ").strip().lower()
        # if list_params == 'y':
        generator.list_parameters()
    except Exception as e:
        print(f"配置表加载失败: {str(e)}")
        sys.exit(1)
    # 交互式指令生成
    print("\n输入指令生成模式 (输入 'q' 退出)")

    while True:
        try:
            # 获取参数名
            param_name = input("\n请输入参数名: ").strip()
            if not param_name:
                continue
            if param_name.lower() in ['q', 'quit', 'exit']:
                print("退出程序")
                break
            # 获取参数值
            param_value = input("请输入参数值: ").strip()
            if not param_value:
                print("参数值不能为空")
                continue
            # 生成指令（自动识别数据类型）
            cmd = generator.generate_command(param_name, param_value)
            print(f"生成写入指令: {cmd[0]}")
            print(f"对应读取指令: {cmd[1]}")

        except ValueError as e:
            print(f"错误: {str(e)}")
        except Exception as e:
            print(f"发生错误: {str(e)}")