import logging
import struct
from tools.log import Log

# 设备Modbus地址表中的参数数据类型，后续用于解码
Date_Type_struct_Map = {
    'uint8_t':'!B',
    'uint16_t':'!H',
    'uint32_t':'!I',
    'float32':'!f',
    'double':'!d'
}

def analysis_message_to_value(memory_value, data_type):
    '''将报文解析到目标数据类型'''
    try:
        if data_type not in Date_Type_struct_Map.keys():
            logging.error('当前数据类型不支持，需维护代码')

        if data_type in ['uint8_t']:
            value_measu = struct.unpack(Date_Type_struct_Map[data_type], bytes(memory_value))[0]
            return value_measu

        bytes_value = []
        for value in memory_value:
            higt_byte = (value & 0xff00) >> 8
            low_byte = (value & 0xff)
            bytes_value.extend([higt_byte, low_byte])
        value_measu = struct.unpack(Date_Type_struct_Map[data_type], bytes(bytes_value))[0]
        return value_measu

    except ValueError as ve:
        logging.error(f"数据处理错误: {ve}")
    except KeyError as ke:
        logging.error(f"字典查找错误: {ke}")
    except struct.error as se:
        logging.error(f"结构体解析错误: {se}")
    except Exception as e:
        logging.error(f"未知错误: {e}")

def analysis_value_to_bytelist(value, datatype):
    """将不同数据类型的参数编码为Modbus 报文要求的参数"""
    try:
        if datatype not in Date_Type_struct_Map:
            logging.error('当前数据类型不支持，需维护代码')
            return False
        if datatype in ['uint8_t', 'uint16_t', 'uint32_t']:
            value = int(value)

        elif datatype in ['float32', 'double']:
            value = float(value)

        byte_data = struct.pack(Date_Type_struct_Map[datatype], value)
        registers = []
        for i in range(0, len(byte_data), 2):
            registers.append((byte_data[i] << 8) + byte_data[i + 1])
        return registers
    except ValueError as ve:
        logging.error(f"数据处理错误: {ve}")
    except KeyError as ke:
        logging.error(f"字典查找错误: {ke}")
    except struct.error as se:
        logging.error(f"结构体解析错误: {se}")
    except Exception as e:
        logging.error(f"未知错误: {e}")

def split_uint16_to_uint8(value):
    high_byte = (value >> 8) & 0xFF
    low_byte = value & 0xFF

    result = [high_byte, low_byte]
    return result
