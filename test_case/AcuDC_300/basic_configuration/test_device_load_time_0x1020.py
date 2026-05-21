import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

device_load_time = [0, 600, 1800, 3600, 0]


def setup_function():
    pass


def test_device_load_time_0x1020():
    rtu_result = []
    for i in device_load_time:
        ret = set_device_load_time(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, True]


def teardown_function():
    pass
