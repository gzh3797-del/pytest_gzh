import logging
from pymodbus.client import ModbusSerialClient, ModbusTcpClient
from modbus_config import modbus_config
import socket
import struct


class ModbusRtuOrTcp:
    def __init__(self, conn_mode='rtu'):
        if modbus_config['conn_mode'] == 'rtu':
            self.client = ModbusSerialClient(port=modbus_config['rtu']['port'],
                                             baudrate=modbus_config['rtu']['baudrate'],
                                             parity=modbus_config['rtu']['parity'])
            self.client.inter_byte_timeout = 100
        elif modbus_config['conn_mode'] == 'tcp':
            self.client = ModbusTcpClient(host=modbus_config['tcp']['ip'], port=modbus_config['tcp']['port'])
        else:
            logging.error('client not exits')
        try:
            self.client.connect()
        except TimeoutError:
            logging.error("modbus rtu connect fail")
            self.client.close()
        # else:
        #     print('modbus rtu connect no error')
        # finally:
        #     print('modbus rtu connect execute completed')

    def close(self):
        self.client.close()

    def write_registers(self, address, values, slave=0):
        try:
            resp = self.client.write_registers(address=address, values=values, slave=slave)
            return resp
        except Exception as e:
            return e

    def read_measurement(self, address, count, slave):
        try:
            resp = self.client.read_holding_registers(address=address, count=count, slave=slave)
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
