import logging
import math
import struct
import time
import openpyxl
from datetime import timedelta

from Config.IOM.modbus_connet import ModbusRtuOrTcp
from Config.IOM.modbus_get_attr import get_single_ai_y_measurement, excel_append_ai_measurement, \
    get_all_ai_y_measurements
from Source.CL3021.source_control import set_dc, close_dc, close_dc_all, read_dc

client = ModbusRtuOrTcp()


def current_time():
    """获取当前时间"""
    time_now = time.time()
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


def res_is_error(res, address):
    """
    判断写入是否成功
    :param res: 写入结果
    :param address: 写入地址
    :return:
    """
    if hasattr(res, 'isError') and res.isError():
        print(f"{current_time()} 警告：写入地址 0x{address:X} 失败")


def float_to_uint32t_4bytes(value):
    """
    将数值（整数或浮点数）转换为：32位无符号整形即两个16位整数，存入两个寄存器（大端序）
    :param value: 要转换的数值（整数或浮点数）
    :return: (high_register, low_register) - 包含两个16进制整数的元组
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


def string_to_uint8t_4bytes(value):
    """
    将中英文字符串转换为：4个8位无符号整形数组，存入两个寄存器（大端序，GBK编码）
    :param value: 要转换的中英文字符串.
    :return: 2个16进制整数的列表（如[0xA1E3, 0x4300]）
    """
    try:
        # 使用GBK编码将字符串转换为字节
        gbk_bytes = value.encode('gbk')

        # 创建一个4字节的缓冲区
        result_bytes = bytearray(4)
        # 复制GBK编码的字节到缓冲区，保持大端序
        for i in range(min(len(gbk_bytes), 4)):
            result_bytes[i] = gbk_bytes[i]

        # 将4字节转换为2个16进制整数（大端序）
        # 第一个整数：前两个字节（高字节在前）
        int1 = (result_bytes[0] << 8) | result_bytes[1]
        # 第二个整数：后两个字节（高字节在前）
        int2 = (result_bytes[2] << 8) | result_bytes[3]

        return [int1, int2]
    except Exception as e:
        logging.error(f"字符串转换为2个16进制整数失败: {e}")
        # 发生错误时返回两个0
        return [0, 0]


def set_ai_param(ai_num, type_line_value, parameter_values):
    """
    修改指定AI口配套参数
    :param ai_num: 1-16
    :param type_line_value: ai_type, line_number
    :param parameter_values: top_limit, bot_limit, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    :return:
    """
    # parameter_keys = ['top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4']
    type_line_address = [0x3000, 0x3005]
    parameter_address = [0x3001, 0x3003, 0x3006, 0x300e, 0x3008, 0x3010, 0x300a, 0x3012, 0x300c, 0x3014]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    print(f"{current_time()} 开始修改AI{ai_num}口的所有配置")
    print(f"'ai_type','line_number'为：{type_line_value}")
    print(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    for addr, value in zip(type_line_address, type_line_value):
        response = client.write_registers(addr + 22 * (ai_num - 1), [value], slave=1)
        res_is_error(response, addr + 22 * (ai_num - 1))
    # 修改所有配套参数
    for address, convert_value in zip(parameter_address, convert_values):
        response = client.write_registers(address + 22 * (ai_num - 1), convert_value, slave=1)
        res_is_error(response, address + 22 * (ai_num - 1))
    client.close()


def set_all_ai_param(type_line_value, parameter_values):
    """
    修改所有AI口配套参数
    :param type_line_value: ai_type, line_number
    :param parameter_values: top_limit, bot_limit, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    :return:
    """
    # 关闭客户端连接
    client.close()
    # 定义AI类型和线路号的寄存器地址
    type_line_address = [0x3000, 0x3005]
    # 定义各个参数的寄存器地址
    parameter_address = [0x3001, 0x3003, 0x3006, 0x300e, 0x3008, 0x3010, 0x300a, 0x3012, 0x300c, 0x3014]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    print(f"{current_time()} 开始修改所有AI口的所有配置")
    print(f"'ai_type','line_number'为：{type_line_value}")
    print(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    for n in range(16):
        for addr, value in zip(type_line_address, type_line_value):
            response = client.write_registers(addr + 22 * n, [value], slave=1)
            # 检查是否写入成功（pymodbus通常用isError()方法检查）
            res_is_error(response, addr + 22 * (n - 1))
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            response = client.write_registers(address + 22 * n, convert_value, slave=1)
            res_is_error(response, address + 22 * (n - 1))
    client.close()


def set_ao_param(ao_num, type_line_value, parameter_values):
    """
    修改指定AO口配套参数
    :param ao_num: 1-16
    :param type_line_value: ao_type, line_number
    :param parameter_values: top_limit, bot_limit, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    :return:
    """
    # parameter_keys = ['top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4']
    type_line_address = [0x3400, 0x3405]
    parameter_address = [0x3401, 0x3403, 0x3406, 0x340e, 0x3408, 0x3410, 0x340a, 0x3412, 0x340c, 0x3414]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    print(f"{current_time()} 开始修改AO{ao_num}口的所有配置")
    print(f"'ao_type','line_number'为：{type_line_value}")
    print(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    for addr, value in zip(type_line_address, type_line_value):
        response = client.write_registers(addr + 22 * (ao_num - 1), [value], slave=1)
        res_is_error(response, addr + 22 * (ao_num - 1))
    # 修改所有配套参数
    for address, convert_value in zip(parameter_address, convert_values):
        response = client.write_registers(address + 22 * (ao_num - 1), convert_value, slave=1)
        res_is_error(response, address + 22 * (ao_num - 1))
    client.close()


def set_all_ao_param(type_line_value, parameter_values):
    """
    修改所有AO口配置参数
    :param type_line_value:
    :param parameter_values:
    :return:
    """
    type_line_address = [0x3400, 0x3405]
    parameter_address = [0x3401, 0x3403, 0x3406, 0x340e, 0x3408, 0x3410, 0x340a, 0x3412, 0x340c, 0x3414]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    print(f"{current_time()} 开始修改所有AO口的所有配置")
    print(f"'ao_type','line_number'为：{type_line_value}")
    print(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    print(type_line_value, parameter_values)
    for n in range(4):
        for addr, value in zip(type_line_address, type_line_value):
            response = client.write_registers(addr + 22 * n, [value], slave=1)
            res_is_error(response, addr + 22 * n)
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            response = client.write_registers(address + 22 * n, convert_value, slave=1)
            res_is_error(response, address + 22 * n)
    client.close()


def set_ao_pmi(ao_num, value):
    """
    配置AO physical measurement Input
    :param ao_num:
    :param value:
    :return:
    """
    address = 0x3950 + 2 * (ao_num - 1)
    rel_value = float_to_uint32t_4bytes(value)
    response = client.write_registers(address, rel_value, slave=1)
    res_is_error(response, address)
    client.close()


def set_all_unit(unit):
    """
    配置所有单位
    :param unit: 2个中文，4个字母或者所有可以输入的特殊字符（"°C"中的"°"：英文状态下：ALT+0176）
    :return:
    """
    value = string_to_uint8t_4bytes(unit)
    for n in range(20):
        if n <= 15:
            address = 0x3200 + 2 * n
            response = client.write_registers(address, value, slave=1)
            res_is_error(response, address)
        else:
            address = 0x34A0 + 2 * (n-16)
            response = client.write_registers(address, value, slave=1)
            res_is_error(response, address)
    client.close()


def iom_test(ai_number=None, ao_number=None, ai_current=None, ai_voltage=None, ao_current=None, ao_voltage=None, expected=None,
             write_to_file=False):
    """
    :param ai_number: 输入通道号
    :param ao_number: 输出通道号
    :param ai_current: 输入电流
    :param ai_voltage: 输入电压
    :param ao_current: 输出电流
    :param ao_voltage: 输出电压
    :param expected: 预期值
    :param write_to_file: 是否写入表格：True/False
    :return:
    """
    ai_start, ai_end = 1, 17
    ao_start, ao_end = 1, 5
    if ai_number or ao_number is None:
        pass
    elif ai_number:
        ai_start, ai_end = ai_number, ai_number+1
    elif ao_number:
        ao_start, ao_end = ao_number, ao_number+1
    if ao_voltage:
        for t in range(ao_start, ao_end):
            if t != 1:
                time.sleep(5)
            print(f"*************************开始执行AO{t}*************************")
            ao_number = t
            for i in range(len(ao_voltage)):
                set_ao_pmi(ao_number, ao_voltage[i])
                time.sleep(5)
                measurement_data = read_dc(0)
                print(
                    f"{i + 1}、现在执行AO {ao_number}，输入电压为{ao_voltage[i]}V，物理测量值为{measurement_data}，预期范围在{expected[i]}, ".replace(" ",""),
                    f"判定结果为：{excel_append_ai_measurement(ao_number, ao_voltage[i], measurement_data, expected[i], write_to_file)}")
            if ao_number == 4:
                print(f"*********************************所有AO口测试结束！*********************************")
            else:
                print(
                    f"*********************************AO{ao_number} 测试完成,你有5s时间切换到AO{ao_number + 1}*********************************")
    elif ao_current:
        for t in range(ao_start, ao_end):
            if t != 1:
                time.sleep(5)
            print(f"*********************************开始执行AO{t}*********************************")
            ao_number = t
            for i in range(len(ao_current)):
                set_ao_pmi(ao_number, ao_current[i])
                time.sleep(5)
                measurement_data = read_dc(1)
                print(
                    f"{i + 1}、现在执行AO {ao_number}，输入电流为{ao_current[i]}V，物理测量值为{measurement_data}，预期范围在{expected[i]}, ".replace(" ",""),
                    f"判定结果为：{excel_append_ai_measurement(ao_number, ao_current[i], measurement_data, expected[i], write_to_file)}")
            if ao_number == 4:
                print(f"*********************************所有AO口测试结束！*********************************")
            else:
                print(
                    f"*********************************AO{ao_number} 测试完成,你有5s时间切换到AO{ao_number + 1}*********************************")
    elif ai_current:
        print("此时AI_Type为电流档,测试所有AI口")
        for ai_number in range(ai_start, ai_end):
        # for ai_number in range(end - 1, start - 1, -1):
            if ai_number != 1:
                time.sleep(3)
            print(f"AI{ai_number}测试开始")
            for n in range(len(ai_current)):
                set_dc(0, ai_current[n])
                time.sleep(6.7)
                measurement_data = get_single_ai_y_measurement(ai_number, client)
                print(
                    f"{n + 1}、现在执行AI {ai_number}，输入电流为{ai_current[n]}mA，物理测量值为{measurement_data}，预期范围在{expected[n]}, ".replace(" ",""),
                    f"判定结果为：{excel_append_ai_measurement(ai_number, ai_current[n], measurement_data, expected[n], write_to_file)}")
            close_dc(2)
            if ai_number == 16:
                print(f"*********************************所有AI口测试结束！*********************************")
            else:
                print(
                    f"*********************************AI{ai_number} 测试完成,你有5s时间切换到AI{ai_number + 1}*********************************")
                time.sleep(2)
    elif ai_voltage:
        print("此时AI_Type为电压档,测试所有AI口")
        for nv in range(len(ai_voltage)):
            if nv != 0:
                time.sleep(4)
            print(f"*********************************测试输入{ai_voltage[nv]}V*********************************")
            set_dc(ai_voltage[nv], 0)
            time.sleep(5)
            measurement_datas = get_all_ai_y_measurements(client)
            time.sleep(0.5)
            for n in range(ai_start, ai_end):  # 循环16个ai口
                measurement = measurement_datas[f"AI{n}"]
                print(f"{n + 1}、AI{n}口输入电压{ai_voltage[nv]}V，物理测量值为{measurement}，预期范围在{expected[nv]}, ".replace(" ",""),
                      f"判定结果为：{excel_append_ai_measurement(n, ai_voltage[nv], measurement, expected[nv], write_to_file)}")
        close_dc_all()


if __name__ == "__main__":
    st = time.time()
    try:
        voltage = [-13,-10,-4,0,2,10,13,14,18,19, 20.5]
        expected = [
            "5.9000 	~6.1000    ",  # 1
            "5.9000 	~6.1000   ",  # 2
            "5.9000 	~6.1000   ",  # 3
            "6.5667 	~6.7667   ",  # 4
            "6.9000 	~7.1000   ",  # 5
            "14.1727 	~14.3727   ",  # 6
            "16.9000 	~17.1000    ",  # 7
            "17.2333 	~17.4333    ",  # 8
            "17.900 	~18.100   ",  # 9
            "17.900 	~18.100",  # 10
            "17.900 	~	18.100",  # 11
            "2.900~3.100      ",  # 12
            "48.000~52.000    ",  # 13
            "48.000~52.000    ",  # 14
            "48.000~52.000    ",  # 15
        ]
        set_all_unit("°C")
        # set_ao_param(1,[0, 3],[10,0,-10,1,10,5,15,7,19,8])
        # set_ai_param(1,[2,3],[20,0,1,-20,5,10,7,15,8,20])
        # iom_test(ao_number=4, ao_current=voltage, expected=expected)
    except KeyboardInterrupt:
        print("\n程序被用户中断，关闭电源输出")
    except Exception as e:
        print("程序异常终止，关闭电源输出")
        logging.error(e)
    finally:
        # close_dc_all()
        et = time.time()
        delta = timedelta(seconds=et - st)
        formatted_time = str(delta).split('.')[0]
        print(f"测试耗时{formatted_time}")

    # set_dc(0, 5)
    # while True:
    #     measurement_data = get_single_ai_y_measurement(1, client)
    #     print(time.time(), measurement_data)
