import logging
import math
import struct
import time

from Config.IOM.modbus_connet import ModbusRtuOrTcp

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


def set_all_ai_type(di_type):
    """
    修改所有AI口ai_type，以及配套参数
    :param di_type: 0-3 : 0-10v, 2-10v, 0-20ma, 4-20ma
    :return:
    """
    parameter_keys = ['top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4']
    parameter_address = [0x3001, 0x3003, 0x3006, 0x300e, 0x3008, 0x3010, 0x300a, 0x3012, 0x300c, 0x3014]
    parameter_values = []
    match di_type:
        case 0:
            parameter_values = [10, 0, 1, 1, 2, 2, 3, 3, 4, 4]
        case 1:
            parameter_values = [10, 2, 2, 2, 3, 3, 4, 4, 5, 5]
        case 2:
            parameter_values = [20, 0, 1, 1, 2, 2, 3, 3, 4, 4]
        case 3:
            parameter_values = [20, 4, 4, 4, 5, 5, 6, 6, 7, 7]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [convert_to_32int_registers(value) for value in parameter_values]
    for n in range(16):
        # 修改所有di_type
        di_type_address = 0x3000 + 22 * n
        print(f"修改ai{n+1}")
        client.write_registers(di_type_address, di_type, slave=1)
        time.sleep(0.3)
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            client.write_registers(address + 22 * n, convert_value, slave=1)
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
    for n in range(16):
        logging.info(f"正在修改ai{n+1}的参数")
        print(f"正在修改ai{n + 1}的参数")
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            client.write_registers(address + 22 * n, convert_value, slave=1)
            time.sleep(0.3)
    client.close()


def set_all_ai_top_bot(parameter_values):
    """
    修改所有AI口top_limit
    :param parameter_values:
    :return:
    """
    # parameter_keys = ['top_limit', 'bot_limit']
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


def set_all_ao_top_bot(parameter_values):
    """
    修改所有AI口top_limit
    :param parameter_values:
    :return:
    """
    # parameter_keys = ['top_limit', 'bot_limit']
    parameter_address = [0x3401, 0x3403]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [convert_to_32int_registers(value) for value in parameter_values]
    for n in range(4):
        print(f"正在修改AO{n + 1}的top和bot为{parameter_values}")
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
    # set_all_ai_type(0)
    # set_all_ai_param([0.7, 0.5, 1, -10, 4, 20, 7, 40, 8, 60])

    # set_all_ao_top_bot([6, 0.2])

    a = [-12,-10, 5,15,20,25, 30, 35, 42, 45]
    b = [
        # "                  ",
        # "  ",
        "4.950 	~	5.050   ",
        "4.950 	~	5.050   ",
        "4.950 	~	5.050   ",
        "4.950 	~	5.050   ",
        "4.950 	~	5.050   ",
        "4.950 	~	5.050   ",
        "5.450 	~	5.550    ",
        "6.200 	~	6.300   ",
        "7.050 	~	7.150    ",
        " 7.2~7.3"
        ]

    st = time.time()
    for t in range(4):
        time.sleep(6)
        print(f"*************************开始执行AO{t+1}*************************")
        ao_number = t+1
        for i in range(len(a)):
            set_ao_pmi(ao_number, a[i])
            print(f"AO{ao_number} 开始配置为{a[i]},预期为{b[i]}V")
            time.sleep(6)
        if ao_number < 4:
            print(f"*************************AO{ao_number} 配置完成,你有6s时间切换到AO{ao_number+1}*************************")
        else:
            print(f"*************************AO{ao_number} 配置完成,测试结束*************************")
    et = time.time()
    print(f"配置耗时{et-st}")

    # set_all_ao_top_bot([7.5, 5])
