from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

network_reset_data = [0, 1, 2]


def setup_function():
    pass


def test_network_reset_0x2007():
    assert network_reset(conn_mode=modbus_config["conn_mode"], value=network_reset_data[0]) is True
    assert network_reset(conn_mode=modbus_config["conn_mode"], value=network_reset_data[1]) is True
    assert network_reset(conn_mode=modbus_config["conn_mode"], value=network_reset_data[2]) is False


def teardown_function():
    pass
