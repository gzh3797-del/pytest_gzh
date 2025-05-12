import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

cable_resistance = [0, 1, 10000, 65535, 65536]


def setup_function():
    pass


def test_cable_resistance_0x1023():
    rtu_result = []
    for i in cable_resistance:
        ret = set_cable_resistance(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, False]


def teardown_function():
    assert set_cable_resistance(conn_mode=modbus_config["conn_mode"], value=0) is True
