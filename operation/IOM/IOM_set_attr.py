import logging
import math
import struct
import time
from api.modbus_connet import ModbusRtuOrTcp
from config.modbus_config import modbus_config
from common.Source.CL3021.source_control import *
from operation.IOM.IOM_get_attr import get_single_ai_y_measurement, excel_append_ai_measurement, \
    get_all_ai_y_measurements


def current_time():
    """获取当前时间"""
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


def res_is_error(res, address):
    """
    判断写入是否成功
    :param res: 写入结果
    :param address: 写入地址
    :return:
    """
    if hasattr(res, 'isError') and res.isError():
        logging.info(f"{current_time()} 警告：写入地址 0x{address:X} 失败")


def float_to_uint32t_4bytes(value, scaling_factor=1000):
    """
    将数值（int32整数或unit32浮点数）转换为：32位无符号整形即两个16位整数，存入两个寄存器（大端序）
    :param scaling_factor: 缩放因子，默认1000
    :param value: 要转换的数值（整数或浮点数）
    :return: (high_register, low_register) - 包含两个16进制整数的元组
    """
    # 应用缩放
    scaled_value = value * scaling_factor
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
        # 复制最多4个字节到缓冲区，保持大端序
        for i in range(min(len(gbk_bytes), 4)):
            result_bytes[i] = gbk_bytes[i]

        # 将4字节转换为2个16进制整数（大端序）
        result_int = struct.unpack('>HH', result_bytes)

        return result_int
    except Exception as e:
        logging.error(f"字符串转换为2个16进制整数失败: {e}")
        return [0, 0]


def set_ai_param(client: ModbusRtuOrTcp, ai_num, type_line_value, parameter_values, slave_id=modbus_config['rtu']['slaveid']):
    """
    修改指定AI口配套参数
    :param client:
    :param slave_id:
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
        response = client.write_registers(addr + 22 * (ai_num - 1), [value], slave=slave_id)
        res_is_error(response, addr + 22 * (ai_num - 1))
    # 修改所有配套参数
    for address, convert_value in zip(parameter_address, convert_values):
        response = client.write_registers(address + 22 * (ai_num - 1), convert_value, slave=slave_id)
        res_is_error(response, address + 22 * (ai_num - 1))
    logging.info(f"AI{ai_num} 修改完成！")
    client.close()


def set_all_ai_param(client: ModbusRtuOrTcp, type_line_value, parameter_values, ai_ao_number=None, slave_id=modbus_config['rtu']['slaveid']):
    """
    修改所有AI口配套参数
    :param client: modbus客户端
    :param slave_id:
    :param ai_ao_number:
    :param type_line_value: ai_type, line_number
    :param parameter_values: top_limit, bot_limit, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    :return:
    """
    client.close()
    type_line_address = [0x3000, 0x3005]
    parameter_address = [0x3001, 0x3003, 0x3006, 0x300e, 0x3008, 0x3010, 0x300a, 0x3012, 0x300c, 0x3014]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    logging.info(f"{current_time()} 开始修改所有AI口的所有配置")
    print(f"{current_time()} 开始修改所有AI口的所有配置为：{type_line_value}，{parameter_values}")
    logging.info(f"'ai_type','line_number'为：{type_line_value}")
    logging.info(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    start, end = 1, 17
    if ai_ao_number:
        start, end = ai_ao_number[0], ai_ao_number[1] + 1
    for n in range(start, end):
        for addr, value in zip(type_line_address, type_line_value):
            response = client.write_registers(addr + 22 * (n-1), [value], slave=slave_id)
            # 检查是否写入成功（pymodbus通常用isError()方法检查）
            res_is_error(response, addr + 22 * (n - 1))
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            response = client.write_registers(address + 22 * (n - 1), convert_value, slave=slave_id)
            res_is_error(response, address + 22 * (n - 1))
        logging.info(f"AI{n} 修改完成！")
        print(f"AI{n} 修改完成！")
    client.close()


def set_ao_param(client: ModbusRtuOrTcp, ao_num, type_line_value, parameter_values, slave_id=modbus_config['rtu']['slaveid']):
    """
    修改指定AO口配套参数
    :param slave_id:
    :param client: modbus客户端
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
    logging.info(f"{current_time()} 开始修改AO{ao_num}口的所有配置")
    logging.info(f"'ao_type','line_number'为：{type_line_value}")
    logging.info(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    for addr, value in zip(type_line_address, type_line_value):
        response = client.write_registers(addr + 22 * (ao_num - 1), [value], slave=slave_id)
        res_is_error(response, addr + 22 * (ao_num - 1))
    # 修改所有配套参数
    for address, convert_value in zip(parameter_address, convert_values):
        response = client.write_registers(address + 22 * (ao_num - 1), convert_value, slave=slave_id)
        res_is_error(response, address + 22 * (ao_num - 1))
    logging.info(f"AO{ao_num} 修改完成！")
    client.close()


def set_all_ao_param(client: ModbusRtuOrTcp, type_line_value, parameter_values, ai_ao_number=None, slave_id=modbus_config['rtu']['slaveid']):
    """
    修改所有AO口配置参数
    :param ai_ao_number: 1-4,不输入时默认所有口
    :param client: modbus客户端
    :param slave_id:
    :param type_line_value: ao_type, line_number
    :param parameter_values: top_limit, bot_limit, X1, Y1, X2, Y2, X3, Y3, X4, Y4
    :return:
    """
    type_line_address = [0x3400, 0x3405]
    parameter_address = [0x3401, 0x3403, 0x3406, 0x340e, 0x3408, 0x3410, 0x340a, 0x3412, 0x340c, 0x3414]
    # 提前将parameter_values转换，以便循环中写入
    convert_values = [float_to_uint32t_4bytes(value) for value in parameter_values]
    logging.info(f"{current_time()} 开始修改所有AO口的所有配置")
    logging.info(f"'ao_type','line_number'为：{type_line_value}")
    logging.info(
        f"'top_limit', 'bot_limit', 'X1', 'Y1', 'X2', 'Y2', 'X3', 'Y3', 'X4', 'Y4'参数为：{parameter_values}")
    start, end = 1, 5
    if ai_ao_number:
        start, end = ai_ao_number[0], ai_ao_number[1] + 1
    for n in range(start, end):
        for addr, value in zip(type_line_address, type_line_value):
            response = client.write_registers(addr + 22 * (n - 1), [value], slave=slave_id)
            res_is_error(response, addr + 22 * (n - 1))
        # 修改所有配套参数
        for address, convert_value in zip(parameter_address, convert_values):
            response = client.write_registers(address + 22 * (n - 1), convert_value, slave=slave_id)
            res_is_error(response, address + 22 * (n - 1))
        logging.info(f"AO{n} 修改完成！")
    client.close()


def set_ao_pmi(client: ModbusRtuOrTcp, ao_num, value, slave_id=modbus_config['rtu']['slaveid']):
    """
    配置AO physical measurement Input
    :param client: modbus客户端
    :param ao_num:
    :param value:
    :return:
    """
    address = 0x3950 + 2 * (ao_num - 1)
    rel_value = float_to_uint32t_4bytes(value)
    response = client.write_registers(address, rel_value, slave=slave_id)
    res_is_error(response, address)
    client.close()


def set_all_ai_unit(client: ModbusRtuOrTcp, unit, slave_id=modbus_config['rtu']['slaveid']):
    """
    配置AIAO板所有单位
    :param client: modbus客户端
    :param slave_id:
    :param unit: 2个中文，4个字母或者所有可以输入的特殊字符（"°C"中的"°"：英文状态下：ALT+0176）
    :return:
    """
    value = string_to_uint8t_4bytes(unit)
    for n in range(20):
        if n <= 15:
            address = 0x3200 + 2 * n
            response = client.write_registers(address, value, slave=slave_id)
            res_is_error(response, address)
        else:
            address = 0x34A0 + 2 * (n-16)
            response = client.write_registers(address, value, slave=slave_id)
            res_is_error(response, address)
    client.close()


def iom_test(client: ModbusRtuOrTcp, ai_number=None, ao_number=None, ai_current=None, ai_voltage=None, ao_current=None, ao_voltage=None, expected=None,
             write_to_file=False):
    """
     IOM自动化用例
    :param ai_number: 输入通道号,格式为[start, end]
    :param ao_number: 输出通道号,格式为[start, end]
    :param client: modbus客户端
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
    judgment_result = True
    if ai_number is not None:
        ai_start, ai_end = ai_number[0], ai_number[1]+1
    if ao_number is not None:
        ao_start, ao_end = ao_number[0], ao_number[1]+1
    if ao_voltage:
        logging.info("此时AI_Type为电压档,测试所有AO口")
        for t in range(ao_start, ao_end):
            if t != 1:
                time.sleep(5)
            logging.info(f"****************************************开始执行AO{t}****************************************")
            ao_number = t
            for i in range(len(ao_voltage)):
                set_ao_pmi(client, ao_number, ao_voltage[i])
                time.sleep(5)
                measurement_data = read_dc(0)
                result = excel_append_ai_measurement(ao_number, ao_voltage[i], measurement_data, expected[i], write_to_file)
                logging.info("AO%d口输入物理测量值为%f，输出电压为%fV，预期范围在%s, 判定结果为：%s",
                               ao_number, ao_voltage[i], measurement_data, expected[i], result)
                if result == "不合格":
                    judgment_result = False
            if ao_number == 4:
                logging.info(f"************************************************所有AO口测试结束！************************************************")
            else:
                logging.info(
                    f"************************************************AO{ao_number} 测试完成,你有5s时间切换到AO{ao_number + 1}************************************************")
    elif ao_current:
        logging.info("此时AI_Type为电流档,测试所有AO口")
        for t in range(ao_start, ao_end):
            if t != 1:
                time.sleep(5)
            print(f"************************************************开始执行AO{t}************************************************")
            ao_number = t
            for i in range(len(ao_current)):
                set_ao_pmi(client, ao_number, ao_current[i])
                time.sleep(5.3)
                measurement_data = read_dc(1)
                result = excel_append_ai_measurement(ao_number, ao_current[i], measurement_data, expected[i], write_to_file)
                logging.info("AO%d口输入物理测量值为%f，输出电流为%fmA，预期范围在%s, 判定结果为：%s",
                             ao_number, ao_current[i], measurement_data, expected[i], result)
                if result == "不合格":
                    judgment_result = False
            if ao_number == 4:
                logging.info(f"************************************************所有AO口测试结束！************************************************")
            else:
                logging.info(
                    f"************************************************AO{ao_number} 测试完成,你有5s时间切换到AO{ao_number + 1}************************************************")
    elif ai_current:
        logging.info("此时AI_Type为电流档,测试所有AI口")
        for ai_number in range(ai_start, ai_end):
        # for ai_number in range(end - 1, start - 1, -1):
            if ai_number != 1:
                time.sleep(3)
            logging.info(f"AI{ai_number}测试开始")
            for n in range(len(ai_current)):
                set_dc(0, ai_current[n])
                time.sleep(6.8)
                measurement_data = get_single_ai_y_measurement(ai_number, client)
                result = excel_append_ai_measurement(ai_number, ai_current[n], measurement_data, expected[n], write_to_file)
                logging.info("AI%d口输入电流%fmA，物理测量值为%f，预期范围在%s, 判定结果为：%s",
                             ai_number, ai_current[n], measurement_data, expected[n], result)
                if result == "不合格":
                    judgment_result = False
            close_dc(2)
            if ai_number == 16:
                logging.info(f"************************************************所有AI口测试结束！************************************************")
            else:
                logging.info(
                    f"************************************************AI{ai_number} 测试完成,你有5s时间切换到AI{ai_number + 1}************************************************")
                time.sleep(2)
    elif ai_voltage:
        logging.info("此时AI_Type为电压档,测试所有AI口")
        for nv in range(len(ai_voltage)):
            if nv != 0:
                time.sleep(4)
            logging.info(f"************************************************测试输入{ai_voltage[nv]}V************************************************")
            set_dc(ai_voltage[nv], 0)
            time.sleep(5)
            measurement_datas = get_all_ai_y_measurements(client)
            time.sleep(0.5)
            for n in range(ai_start, ai_end):  # 循环16个ai口
                measurement = measurement_datas[f"AI{n}"]
                result = excel_append_ai_measurement(n, ai_voltage[nv], measurement, expected[nv], write_to_file)
                logging.info("AI%d口输入电压%fV，物理测量值为%f，预期范围在%s, 判定结果为：%s",
                             n, ai_voltage[nv], measurement, expected[nv], result)
                if result == "不合格":
                    judgment_result = False
    close_dc_all()
    return judgment_result


def set_all_di_unit(client: ModbusRtuOrTcp, unit, slave_id=modbus_config['rtu']['slaveid']):
    """
    配置DIDO板所有单位
    :param client: modbus客户端
    :param slave_id:
    :param unit: 2个中文，4个字母或者所有可以输入的特殊字符（"°C"中的"°"：英文状态下：ALT+0176）
    :return:
    """
    # value = [string_to_uint8t_4bytes(unit) for _ in range(28)]
    value = string_to_uint8t_4bytes(unit)
    for n in range(28):
        address = 0x2100 + 2 * n
        response = client.write_registers(address, value, slave=slave_id)
        res_is_error(response, address)
    client.close()


def set_all_di_pulse_constant(client: ModbusRtuOrTcp, value, slave_id=modbus_config['rtu']['slaveid']):
    """
    配置所有di_pulse_constant
    :param slave_id:
    :param client:
    :param value:
    :return:
    """
    rel_value = float_to_uint32t_4bytes(value)
    for n in range(28):
        address = 0x2001 + 3 * n
        response = client.write_registers(address, rel_value, slave=slave_id)
        res_is_error(response, address)
    print(f"所有DI已脉冲常量配置为：{value}")
    client.close()


def set_all_di_pulse_count(client: ModbusRtuOrTcp, value, slave_id=modbus_config['rtu']['slaveid']):
    """
    配置所有di_pulse_count
    :param slave_id:
    :param client:
    :param value:
    :return:
    """
    rel_value = float_to_uint32t_4bytes(value,scaling_factor = 1)
    for n in range(28):
        address = 0x2200 + 2 * n
        response = client.write_registers(address, rel_value, slave=slave_id)
        res_is_error(response, address)
    print(f"所有DI已脉冲常量配置为：{value}")
    client.close()


if __name__ == "__main__":
    client = ModbusRtuOrTcp()
    # st = time.time()
    set_all_di_unit(client, "米")
    # time.sleep(4)
    # set_all_ai_unit(client, "厘米")
    # time.sleep(4)
    # set_all_ai_unit(client, "m")
    # time.sleep(4)
    # set_all_ai_unit(client, "cm")
    # time.sleep(4)
    # set_all_ai_unit(client, "~!@#")
    # time.sleep(4)
    # set_all_ai_unit(client, "$%^*")
    # time.sleep(4)
    # set_all_ai_unit(client, "()_+")
    # time.sleep(4)
    # set_all_ai_unit(client, "摄氏度")
    # time.sleep(4)
    # set_all_ai_unit(client, "ABCDE")
    # time.sleep(4)
    # set_all_ai_unit(client, "°C")
    # time.sleep(4)
    # set_all_di_pulse_count(client, 100)







