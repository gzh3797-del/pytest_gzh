from crcmod import mkCrcFun
from binascii import unhexlify


def crc16_modbus(s):
    crc16 = mkCrcFun(0x18005, rev=True, initCrc=0xFFFF, xorOut=0x0000)
    data = s.replace(' ', '')
    crc_out = hex(crc16(unhexlify(data))).upper()
    str_list = list(crc_out)
    if len(str_list) == 5:
        str_list.insert(2, '0')  # 位数不足补0
    crc_data = ''.join(str_list[2:])
    if crc_data[2:] == "":
        crc_data = "00" + crc_data
    if crc_data[:2] == "":
        crc_data = crc_data + "00"
    return crc_data[2:] + '' + crc_data[:2]
