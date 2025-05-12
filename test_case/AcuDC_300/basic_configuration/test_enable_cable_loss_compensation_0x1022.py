import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

enable_cable_loss_compensation = [0, 1, 2]


def setup_function():
    pass


def test_enable_cable_loss_compensation_0x1022():
    rtu_result = []
    for i in enable_cable_loss_compensation:
        ret = set_cable_loss_compensation_enable(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, False]


def teardown_function():
    assert set_cable_loss_compensation_enable(conn_mode=modbus_config["conn_mode"], value=0) is True
