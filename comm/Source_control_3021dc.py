from comm.source_control import *
import struct

"""
帧格式：Head + RxID + TxID + Length + Cmd + Data + CS
其中：
    Head	帧起始标志81H  一个字节
    RxID	接受ID 一个字节 
    TxID	发送ID 一个字节
    Length 	数据包长度 一个字节
    Cmd 	命令 一个字节
    Data 	数据 Length - 6
    CS		检验码 从RxID到 Data 所有数据的异或和
数据解析:
    1、在表达浮点型数据时，采用 Int4E1 结构体方式。每个数值占用 5 个字节
    2、值以[小端模式]写入或读取
    例如： 
    接收到的数据串 E8 CD 08 00 FC 计算如下： 
    首先这个数据分成两部分：E8 CD 08 00 和 FC 
    其中 E8 CD 08 00 是有符号长整型,转成小端模式理解应该是 0x0008CDE8(即字节顺序调转) 
    0x0008CDE8 转成十进制值即 577000 
    而 0xFC 是有符号字节型，转成十进制值即 -4 
    它们合起来表达的浮点数值是：577000 / 10000 = 57.7
"""

def set_grade_source_dc(mode: int = 3):
    """
    function: 设置直流电流的电压和直流电流的档位
    参数说明:
    自动手动标志(data)：
    1--手动电压
    2--手动电流
    3--电压电流都手动
    """
    set_cmd = [0x81, 0x01, 0x26, 0x09, 0x32, mode, 0x00, 0x00]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def set_measurement_mode_dc(measurement_mode: int = 0):
    """
    function: 设置测量直流表的测量模式
    参数说明(measurement_mode):
        0,同时测量电压和电流
        1,测量直流电压
        2,测量直流电流
    """
    set_cmd = [0x81, 0x01, 0x26, 0x07, 0x3C, measurement_mode]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    if bytes.hex(ret[1][5]).upper() == "0x30":
        logging.error("设置直流电表的测量模式{}".format({0: "同时测量电压和电流", 1: "测量直流电压", 2: "测量直流电流"}.get(measurement_mode)))
        print("设置直流电表的测量模式{}".format({0: "同时测量电压和电流", 1: "测量直流电压", 2: "测量直流电流"}.get(measurement_mode)))
    else:
        logging.error("设置直流电表的测量模式失败")
    source_control.close()

def set_grade_meter_dc(mode: int = 3):
    """
    function: 设置直流电表的电压和直流电流的档位
    参数说明:
    自动手动标志(data)：
    1--手动电压
    2--手动电流
    3--电压电流都手动
    """
    set_cmd = [0x81, 0x01, 0x26, 0x09, 0x34, mode, 0x00, 0x00]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    if bytes.hex(ret[1][5]).upper() == "0x30":
        logging.error("设置直流电表的电压和直流电流的档位: {}. 1--手动电压, 2--手动电流, 3--电压电流都手动".format(mode))
        print("设置直流电表的电压和直流电流的档位: {}. 1--手动电压, 2--手动电流, 3--电压电流都手动".format(mode))
    else:
        logging.error("设置直流源失败")
    source_control.close()

def set_output_dc(power_type: int = 0x03, Voltage_dc: float = 0.000, current_dc: float = 0.000):
    """
    function: 设置直流源电压和直流电流数值
    参数说明:
    输出控制标志(power_type)：
    1—输出电压
    2—输出电流
    3--电压电流都输出
    """
    set_cmd = [0x81, 0x01, 0x26, 0x11, 0x31, power_type]
    pdu = str(hex(int(Voltage_dc * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(current_dc * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    if bytes.hex(ret[1][5]).upper() == "0x30":
        logging.error("设置直流源成功Voltage_DC: {},设置直流源成功current_DC: {}".format(Voltage_dc, current_dc))
        print("设置直流源成功Voltage_DC: {},设置直流源成功current_DC: {}".format(Voltage_dc, current_dc))
    else:
        logging.error("设置直流源失败")
    source_control.close()


def bytes_to_float_little_endian(data_bytes: bytes):
    """
    Function: bytes类型数据转换成float类型数据
    数据说明:
    处理数据格式Int4E1 结构体，有5个字节
    """
    data_hex = bytes.fromhex(str(data_bytes))
    # 提取前4字节为整数（小端）
    mantissa = struct.unpack('<i', data_hex[:4])[0]  # <i 表示小端有符号32位整数
    # 提取指数
    exponent = struct.unpack('b', data_hex[4:])[0]  # b 表示有符号8位整数
    # 计算最终值
    data_float = mantissa * (10 ** exponent)
    return data_float


def close_source_dc(source_type: int):
    """
    标志位：
    BIT0=1,电压关闭
    BIT1=1,电流关闭
    只能 2 选一,其余 BIT 置 0
    """
    set_cmd = [0x81, 0x01, 0x26, 0x07, 0x38, source_type]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    if bytes.hex(ret[1][5]).upper() == "0x30":
        logging.error("关闭直流源成功")
    else:
        logging.error("关闭直流源失败")
    source_control.close()

def read_measurement_dc():
    """
    function: 获取电源测量的电压和电流值
    参数说明:
    读取Cmd 0XA3
    返回失败Cmd 0X33
    返回成功cmd 0X53
    """
    set_cmd = [0x81, 0x01, 0x26, 0x06, 0xa3]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    raw_bytes = ret[1]
    print(f"收到报文(16进制字符串)：{raw_bytes.hex().upper()}")
    if raw_bytes[5].hex().upper() == "0x53":
        if raw_bytes[6].hex().upper() == "0x01":
            logging.error("电压过载，结束测试")
        elif raw_bytes[6].hex().upper() == "0x02":
            logging.error("电流过载，结束测试")
        elif raw_bytes[6].hex().upper() == "0x03":
            logging.error("电压电流过载，结束测试")
        else:
            return {
                  "voltage_ripple": bytes_to_float_little_endian(raw_bytes[7:11]),  # 小端 bytes 数据
                  "current_ripple": bytes_to_float_little_endian(raw_bytes[12:16]),  # 小端 bytes 数据
                  "voltage_amplitude": bytes_to_float_little_endian(raw_bytes[17:21]),  # 小端 bytes 数据
                  "current_amplitude": bytes_to_float_little_endian(raw_bytes[22:26]),  # 小端 bytes 数据
              }
    else:
        print("获取电压电流失败")
    source_control.close()

def clear_overload_dc_lock(overload_flag: int = 3):
    """
    清除直流过载锁定
    :param overload_flag:
             BIT0=0,清除电压过载
             BIT1=0,清除电流过载
            其他 BIT 无效忽略
    """
    set_cmd = [0x81, 0x01, 0x25, 0x07, 0x39, hex(overload_flag)]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    if bytes.hex(ret[1][5]).upper() == "0x39":
        logging.error("清除过载成功")
    else:
        logging.error("清除过载失败")
    source_control.close()

