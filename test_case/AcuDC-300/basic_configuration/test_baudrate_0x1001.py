from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

baudrate = [1200, 2400, 4800, 9600, 19200, 38400, 57600, 7680, 11520, 12800]


def setup_function():
    pass


def test_baudrate_0x1001():
    rtu_baudrate = []
    for i in baudrate:
        ret = set_rs485_baudrate(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_baudrate.append(ret)
    for rtu in range(1, len(rtu_baudrate) - 1):
        assert rtu_baudrate[rtu] is True
    assert rtu_baudrate[-1] is False and rtu_baudrate[0] is False


def teardown_function():
    assert set_rs485_baudrate(conn_mode='tcp', value=19200) is True
