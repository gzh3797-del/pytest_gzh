import logging
from pymodbus.client import ModbusSerialClient, ModbusTcpClient
from Common.modbus_config import modbus_config
import socket
import serial
import struct
import crcmod


class ModbusRtuOrTcp:
    def __init__(self, conn_mode='rtu'):
        """
        通过ModbusSerialClient对连接板子，可以通过串口与板子通信
        :param conn_mode:
        """
        if modbus_config['conn_mode'] == 'rtu':
            self.client = ModbusSerialClient(port=modbus_config['rtu']['port'],
                                             baudrate=modbus_config['rtu']['baudrate'],
                                             parity=modbus_config['rtu']['parity'])
            self.client.inter_byte_timeout = 0.02
            self.client.timeout = 0.5
        elif modbus_config['conn_mode'] == 'tcp':
            self.client = ModbusTcpClient(host=modbus_config['tcp']['ip'], port=modbus_config['tcp']['port'])
        else:
            logging.error('client not exits')
        try:
            self.client.connect()
        except Exception as e:
            logging.error(f"modbus rtu connect fail: {str(e)}")
            if hasattr(self, 'client') and self.client:
                self.client.close()
        else:
            logging.info('modbus rtu connect no error')
        finally:
            logging.info('modbus rtu connect execute completed')

    def close(self):
        self.client.close()

    def write_registers(self, address, values, slave):
        """
        写入多个寄存器
        :param address:
        :param values:
        :param slave:
        :return:
        """
        try:
            resp = self.client.write_registers(address=address, values=values, device_id=slave)
            return resp
        except Exception as e:
            return e

    def write_register(self, address, value, slave):
        """
        写入单个寄存器
        :param address:
        :param value:
        :param slave:
        :return:
        """
        try:
            resp = self.client.write_register(address=address, value=value, device_id=slave)
            return resp
        except Exception as e:
            return e

    def read_measurement(self, address, count, slave):
        """
        读取测量数据，即读取保持寄存器数据
        :param address:
        :param count:
        :param slave:
        :return:
        """
        try:
            resp = self.client.read_holding_registers(address=address, count=count, device_id=slave)
            logging.info('read_measurement ret is:{}'.format(resp))
            if resp.isError():
                return "resp is error"
            measurement = resp.registers
            return measurement
        except Exception as e:
            return e


class ModbusTcp6A:
    def __init__(self, timeout=0.5):
        self.ip = modbus_config['tcp']['ip']
        self.port = modbus_config['tcp']['port']
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.settimeout(timeout)
        self.isSucess = None
        try:
            self.socket.connect((self.ip, self.port))
        except TimeoutError:
            self.isSucess = False
            raise Exception("modbus tcp connect fail")
        else:
            print("modbus tcp connect success")
            self.isSucess = True

    def read_funcode_03(self, addr: int, count: int = 1, slaveid: int = 1):

        _read_byarry = bytearray([
            0x00, 0x01,
            0x00, 0x00,
            0x00, 0x06,
            slaveid,
            0x03,
            addr >> 8 & 0xff,
            addr & 0xff,
            count >> 8 & 0xff,
            count & 0xff
        ])
        print(_read_byarry)
        if self.isSucess:
            try:
                self.socket.send(_read_byarry)
            except TimeoutError:
                self.socket.close()
                raise Exception("mosbus tcp read fail")
            else:
                byte = self.socket.recv(1024)
                self.socket.close()
                return struct.unpack(f">{count}H", bytearray(byte[9:]))

    def write_registers(self, start_addr, values: list, slaveid, funccode=0x6A):
        length = len(values) * 2 + 7
        # 报文头
        bmap = bytearray([
            0x00, 0x01,
            0x00, 0x00,
            length >> 8 & 0xff,  # 写入的字节数量高位
            length & 0xff,  # 写入的字节数量低位
            slaveid  # 单元标识
        ])

        pdu = bytearray(
            [
                funccode,  # 功能码
                start_addr >> 8 & 0xff,  # 起始地址高位字节
                start_addr & 0xff,  # 起始地址低位字节
                len(values) >> 8 & 0xff,  # 写入的数量高字节
                len(values) & 0xff,  # 写入的数量低字节
                len(values) * 2  # 字节数
            ]
        )
        for value in values:
            pdu.extend([(value >> 8) & 0xff, value & 0xff])

        request = bmap + pdu
        if self.isSucess:
            self.socket.send(request)
            data_recv = self.socket.recv(1024)
            self.socket.close()
        else:
            self.socket.close()
            raise Exception("modbus tcp connect fail")
        if funccode == 0x6A:
            return start_addr, len(values), data_recv  # 返回报文
        else:
            return struct.unpack(f">{len(values) * 2}H", bytearray(data_recv[8:])), data_recv


class SerialRtu:
    def __init__(self, conn_mode='rtu'):
        try:
            self.ser = serial.Serial(port=modbus_config['rtu']['port'], baudrate=modbus_config['rtu']['baudrate'],
                                     parity=modbus_config['rtu']['parity'])
            if self.ser.is_open:
                logging.info("打开串口成功。")
            else:
                logging.error("打开串口失败。")
        except Exception as e:
            logging.error("打开串口异常：{}".format(e))

    def close(self):
        self.ser.close()

    def write_func6A_registers(self, address, count, values: list, slave, funccode=0x6A):
        data = [
            slave,
            funccode,  # 功能码
            address >> 8 & 0xff,  # 起始地址高位字节
            address & 0xff,  # 起始地址低位字节
            len(values) >> 8 & 0xff,  # 写入的数量高字节
            len(values) & 0xff,  # 写入的数量低字节
            len(values) * 2  # 字节数
        ]
        pdu = bytearray(data)
        for value in values:
            pdu.extend([(value >> 8) & 0xff, value & 0xff])
        crc32_func = crcmod.mkCrcFun(0x18005, rev=True, initCrc=0xFFFF, xorOut=0x0000)
        ret1 = str(hex(crc32_func(pdu)))
        pdu.extend([int(ret1[4:6], 16) & 0xff])
        pdu.extend([int(ret1[2:4], 16) & 0xff])
        try:
            address_hex = hex(int(address))[2:6]
            if count < 16:
                count_hex = '0' + hex(int(count))[2:4]
            else:
                count_hex = hex(int(count))[2:4]
            if count < 8:
                count_hex_2 = '0' + hex(int(count * 2))[2:4]
            else:
                count_hex_2 = hex(int(count * 2))[2:4]
            compare_str = str('0' + str(slave)) + '6a' + str(address_hex) + '00' + str(count_hex)
            ret1 = self.ser.write(pdu)
            str1 = ''
            for i in range(6):
                byte_seq = self.ser.read()
                byte_seq1 = byte_seq[0]
                str1 += f'0x{byte_seq1:02x}'
            str1 = str1.replace('0x', '')
            if str1 == compare_str:
                return True
            else:
                return False
        except Exception as e:
            return e