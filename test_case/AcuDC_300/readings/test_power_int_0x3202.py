from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_power_int_0x3202():
    assert set_power_client(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_power_client(conn_mode=modbus_config["conn_mode"])) is int


def teardown_function():
    pass
