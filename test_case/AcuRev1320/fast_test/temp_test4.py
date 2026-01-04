import struct


def get_bytes_value(memory_value):
    """
    解析寄存器返回值
    :param memory_value: 寄存器返回值
    :return: 整数列表
    """
    bytes_value = []
    for value in memory_value:
        high_byte = (value & 0xff00) >> 8
        low_byte = (value & 0x00ff)
        bytes_value.extend([high_byte, low_byte])
    return bytes_value


regs = [0x1234, 0xABCD]
bytes_value = get_bytes_value(regs)
value = struct.unpack('!f', bytes(bytes_value))[0]
print(bytes_value)
print(value)
print(bytes(bytes_value))
print(struct.unpack('!f', bytes(bytes_value)))

