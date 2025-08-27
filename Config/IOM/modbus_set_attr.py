import logging
import math
import struct
import time
import openpyxl

from Config.IOM.modbus_connet import ModbusRtuOrTcp
from Config.IOM.modbus_get_attr import get_single_ai_y_measurement, excel_append_ai_measurement, \
    get_all_ai_y_measurements
from Source.CL3021.source_control import set_dc, close_dc, close_dc_all

client = ModbusRtuOrTcp()


def convert_to_32int_registers(value):
    """
    将数值（整数或浮点数）转换为32位整形Modbus寄存器值（大端序）
    :param value:
    :return: (high_register, low_register) - 包含两个16位整数的元组
    """
    # 应用缩放
    scaled_value = value * 1000
    # 检查是否为有限值
    if not math.isfinite(scaled_value):
        raise ValueError("转换后值为非有限数(inf或nan)")
    # 缩放后取整
    if scaled_value >= 0:
        int_value = math.floor(scaled_value + 0.5)  # 正数四舍五入
    else:
        int_value = math.ceil(scaled_value - 0.5)  # 负数四舍五入
    # 范围检查（32位有符号整数）
    if int_value < -2147483648 or int_value > 2147483647:
        raise ValueError(f"转换后值 {int_value} 超出32位整数范围")

    # 大端序打包
    packed_data = struct.pack('>i', int_value)
    # 拆分为两个16位整数（高位在前）
    return struct.unpack('>HH', packed_data)


def set_ai_type(ai_number, value):
    """
    配置指定AI口ai_type
    :param ai_number: 1-16
    :param value: 0-3 : 0-10v, 2-10v, 0-20ma, 4-20ma
    :return:
    """
    address = 0x3000 + 22 * (ai_number-1)
    ret = client.write_registers(address, value, slave=1)
    if '(4117,1)' in str(ret):
        logging.error('set nonsupport set, but CT2 set success, ret is:{}'.format(ret))
        return True
    client.close()
    return False


def set_all_ai_type(ai_type):
    """
    修改所有AI口ai_type
    :param ai_type:  0-3 : 0-10v, 2-10v, 0-20ma, 4-20ma
    :return:
    """
    for n in range(8):
        # 修改所有ai_type
        ai_type_address = 0x3000 + 22 * n
        print(f"修改ai{n+1}")
        client.write_registers(ai_type_address, ai_type, slave=1)
        time.sleep(0.3)
    client.close()


def set_all_ai_param(parameter_values):
    """
    修改所有AI口配套参数
    :return:
    """
    # parameter_keys = ['top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4']
    parameter_address = [0x3001, 0x3003, 0x3006, 0x300e, 0x3008, 0x3010, 0x300a, 0x3012, 0x300c, 0x3014]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [convert_to_32int_registers(value) for value in parameter_values]
    print(f"开始修改配置'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'为：")
    print(parameter_values)
    for n in range(8):
        logging.info(f"正在修改ai{n+1}的参数")
        print(f"正在修改ai{n + 1}的top，bot以及点坐标")
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            client.write_registers(address + 22 * n, convert_value, slave=1)
            time.sleep(0.3)
    client.close()


def set_all_ai_top_bot(parameter_values):
    """
    修改所有AI口top_limit，bot_limit
    :param parameter_values:
    :return:
    """
    parameter_address = [0x3001, 0x3003]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [convert_to_32int_registers(value) for value in parameter_values]
    for n in range(16):
        print(f"正在修改AI{n + 1}的top和bot为{parameter_values}")
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            client.write_registers(address + 22 * n, convert_value, slave=1)
            time.sleep(0.3)
    client.close()


def set_all_ao_param(parameter_values):
    """
    修改所有AO口，top_limit，bot_limit和配置点参数
    :param parameter_values:
    :return:
    """
    parameter_address = [0x3401, 0x3403, 0x3406, 0x340e, 0x3408, 0x3410, 0x340a, 0x3412, 0x340c, 0x3414]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [convert_to_32int_registers(value) for value in parameter_values]
    for n in range(2):
        print(f"正在修改AO{n + 1}的配置为{parameter_values}")
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            client.write_registers(address + 22 * n, convert_value, slave=1)
            time.sleep(0.3)
    client.close()


def set_ao_pmi(ao_number, value):
    """
    配置AO physical measurement Input
    :param ao_number:
    :param value:
    :return:
    """
    address = 0x3950 + 2 * (ao_number - 1)
    rel_value = convert_to_32int_registers(value)
    client.write_registers(address, rel_value, slave=1)
    client.close()


if __name__ == "__main__":
    # set_all_ao_param([19.5, 4.5,-10,5, 2,7, 13,17, 19,19])
    # ao_voltage = []
    ao_voltage = [-20,-13,-10,19, 20.5,25]
    current = []
    voltage = []
    expected = [
        "4.4000 ~ 4.6000       ",  # 1
        "4.4000 ~ 4.6000           ",  # 2
        "4.9000 ~ 5.1000     ",  # 3
        "18.9000 ~ 19.1000           ",  # 4
        "19.4000 ~ 19.6000                   ",  # 5
        "19.4000 ~ 19.6000         ",  # 6
        "17.9000 	~	18.1000              ",  # 7
        "18.9000 ~ 19.1000        ",  # 8
        "7.9500 ~ 8.0500          ",  # 9
        "6.9500 ~ 7.0500          ",  # 10
        "6.9500 ~ 7.0500          ",  # 11

    ]

    st = time.time()
    if ao_voltage:
        for t in range(2):
            if t != 0:
                time.sleep(4)
            print(f"*************************开始执行AO{t + 1}*************************")
            ao_number = t + 1
            for i in range(len(ao_voltage)):
                set_ao_pmi(ao_number, ao_voltage[i])
                print(f"AO{ao_number} 开始配置为{ao_voltage[i]},预期为{expected[i]}V")
                time.sleep(6.5)
            if ao_number < 2:
                print(
                    f"*************************AO{ao_number} 配置完成,你有4s时间切换到AO{ao_number + 1}*************************")
            else:
                print(f"*************************AO{ao_number} 配置完成,测试结束*************************")
    elif current:
        print("此时AI_Type为电流档,测试所有AI口")
        for ai_number in range(1, 16):
            if ai_number != 1:
                time.sleep(3)
            print(f"AI{ai_number}测试开始")
            for n in range(len(current)):
                set_dc(0, current[n])
                time.sleep(7)
                measurement_data = get_single_ai_y_measurement(ai_number,client)
                close_dc(2)
                print(f"现在执行AI {ai_number}，输入电流为{current[n]}mA，实际测量值为{measurement_data}，预期范围在{expected[n]}, "
                             f"判定结果为：{excel_append_ai_measurement(ai_number, 2, current[n], measurement_data, expected[n])}")
            close_dc_all()
            if ai_number == 8:
                print(f"所有AI口测试结束！=====================================================================")
            else:
                print(
                    f"AI {ai_number} 测试结束，你有3s时间切换到AI {ai_number + 1}！=====================================================================")
            close_dc_all()
    elif voltage:
        print("此时AI_Type为电压档,测试所有AI口")
        for nv in range(len(voltage)):
            print(f"测试输入{voltage[nv]}V=====================================================================")
            set_dc(voltage[nv], 0)
            time.sleep(6)
            measurement_datas = get_all_ai_y_measurements(client)
            close_dc(1)
            time.sleep(1)
            for n in range(8):  # 循环16个ai口
                measurement = measurement_datas[f"AI{n+1}"]
                print(f"AI{n+1}口输入电压{voltage[nv]}V，实际测量值为{measurement}，预期范围在{expected[nv]}, "
                      f"判定结果为：{excel_append_ai_measurement(n+1, 0, voltage[nv], measurement, expected[nv])}")
        close_dc_all()
    et = time.time()
    print(f"配置耗时{et - st}")
