from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

clear_max_min_data = [0, 1, 2]


def setup_function():
    pass


def test_clear_max_min_0x2003():
    assert clear_max_min(conn_mode=modbus_config["conn_mode"], value=clear_max_min_data[0]) is True
    assert clear_max_min(conn_mode=modbus_config["conn_mode"], value=clear_max_min_data[1]) is True
    assert clear_max_min(conn_mode=modbus_config["conn_mode"], value=clear_max_min_data[2]) is False


def teardown_function():
    pass
