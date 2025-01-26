import time

from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

tcp_port = [1, 4, 502, 65533, 65534, 65535]


def setup_function():
    pass


def test_tcp_port_0x1011():
    tcp_tcp_port = []
    for i in tcp_port:
        ret = set_modbus_tcp_port(conn_mode=modbus_config["conn_mode"], value=i)
        tcp_tcp_port.append(ret)
    # print(tcp_tcp_port)
    for tcp in range(len(tcp_tcp_port) - 1):
        assert tcp_tcp_port[tcp] is True
    assert tcp_tcp_port[-1] is False


def teardown_function():
    assert set_modbus_tcp_port(conn_mode=modbus_config["conn_mode"], value=502) is True
