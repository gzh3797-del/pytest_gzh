from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_current_int_0x3201():
    assert set_current_client(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_current_client(conn_mode=modbus_config["conn_mode"])) is int


def teardown_function():
    pass
