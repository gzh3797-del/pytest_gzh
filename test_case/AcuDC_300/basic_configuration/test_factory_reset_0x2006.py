import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

factory_reset_data = [0, 1, 2]


def setup_function():
    pass


def test_factory_reset_0x2006():
    assert factory_reset(conn_mode=modbus_config["conn_mode"], value=factory_reset_data[0]) is True
    assert factory_reset(conn_mode=modbus_config["conn_mode"], value=factory_reset_data[1]) is False
    time.sleep(8)
    assert factory_reset(conn_mode=modbus_config["conn_mode"], value=factory_reset_data[2]) is False


def teardown_function():
    pass
