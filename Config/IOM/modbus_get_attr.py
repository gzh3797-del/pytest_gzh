import os
import pandas as pd
from typing import Any
from Config.IOM.modbus_connet import ModbusRtuOrTcp
import struct
from datetime import datetime
import time



def convert_energy_registers(registers_data):
    """
    将寄存器读取到的数据转换为浮点类型。
    （原始数据是 大端序（Big Endian）格式的浮点数，需要采用 小端序（Little Endian）格式存储。
    同时，每组寄存器（2个寄存器4个字节8位16进制）需要按小端序解释为一个32位浮点数。）
    :param registers_data:
    :return:
    """
    results = []
    # 每2个寄存器转换成一个浮点数（2个寄存器4个字节，4个字节是8位16进制）
    for i in range(0, len(registers_data), 2):
        # 提取低位高位寄存器值，每个寄存器2个字节4位16进制
        low_reg = registers_data[i + 1]  # 51149(0xc7cd)
        high_reg = registers_data[i]  # 16543(0x409f)
        # 处理零值
        if low_reg == 0 and high_reg == 0:
            results.append(0.0)
            continue
        # 将寄存器转换为小端序浮点数。
        low_bytes = low_reg.to_bytes(2, 'little')  # c7cd->cdc7
        high_bytes = high_reg.to_bytes(2, 'little')  # 409f->9f40
        # 将两个小端序直接解释为浮点数，c7cd409f->2.999....
        float_val = struct.unpack('<f', low_bytes + high_bytes)[0]
        results.append(float_val)

    return results


def excel_append_ai_measurement(ai_number, ai_type, input_data, measurement, range_str):
    """
    将AI测量数据追加到Excel文件中
    ai_number -- AI口编号 (1-16的整数)
    input_data -- 输入值 (浮点数)
    measurement -- 实测值 (浮点数)
    range_str -- 范围字符串 (如 "-35.4000~-34.600")
    """
    # 检查AI口编号范围
    if not 1 <= ai_number <= 16:
        raise ValueError("AI口编号必须为1-16的整数")
    # 解析范围字符串
    try:
        range_min, range_max = map(float, range_str.strip().split('~'))
    except ValueError:
        raise ValueError("范围格式错误，应类似'-35.4000~-34.600'")
    # 生成当前时间戳
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 创建新数据行
    new_row = {
        "时间": timestamp,
        "AI口": f"AI{ai_number:02d}",
        "AI_TYPE": ai_type,
        "输入值": input_data,
        "实测值": measurement,
        "范围": range_str
    }
    # 生成判定结果
    if range_min <= measurement <= range_max:
        new_row["判定结果"] = "合格"
    else:
        new_row["判定结果"] = "不合格"
    # 文件路径
    file_path = "ai_i_measurements.xlsx"
    # 创建分割行（全空值行）
    separator_row = {col: "" for col in new_row.keys()}  # 所有列值为空字符串[7](@ref)
    # 处理文件存在与否的逻辑
    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        # 合并新数据与分割行
        df_combined = pd.concat([df_existing, pd.DataFrame([new_row]), pd.DataFrame([separator_row])],
                                ignore_index=True)
    else:
        # 首次创建文件时，先添加数据行再添加分割行
        df_combined = pd.concat([pd.DataFrame([new_row]), pd.DataFrame([separator_row])], ignore_index=True)
    df_combined.to_excel(file_path, index=False)
    return new_row["判定结果"]


def get_all_ai_y_measurements(modbus_client: ModbusRtuOrTcp):
    """
    读取所有AI输入
    :return:
    """
    # 列表推导式生成 ['AI1','AI2','AI3',,,,]
    measurement_key_list = [f'AI{i}' for i in range(1, 17)]
    ret = {}
    try:
        # 读取多个寄存器
        energy: list = modbus_client.read_measurement(address=0x3700, count=32, slave=1)
        measurement_value_list = convert_energy_registers(energy)
        # 合并两个列表为字典
        for key, value in zip(measurement_key_list, measurement_value_list):
            ret[key] = value
    except Exception as e:
        print(f"连接失败: {e}")
    return ret


def get_single_ai_y_measurement(ai_number, modbus_client) -> float | None:
    """
    读取单个AI输入
    :param ai_number: 1-16
    :param modbus_client: pytest fixture 提供的 ModbusRtuOrTcp 实例
    :return: 解码后的 float 类型值
    """
    # 生成AI1-AI16对应16个寄存器地址的键值对:{1:0x3500,2:0x3502,3:0x3504,,,,,,}
    measurement_map = {i: 0x3700 + 2 * (i - 1) for i in range(1, 17)}
    try:
        registers = modbus_client.read_measurement(address=measurement_map[ai_number], count=2, slave=1)
        measurement_value = convert_energy_registers(registers)
        return measurement_value[0]
    except Exception as e:
        return None


if __name__ == "__main__":

    get_single_ai_y_measurement(1)
