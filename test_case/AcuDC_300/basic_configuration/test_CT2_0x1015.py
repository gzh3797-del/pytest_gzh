from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

ct2 = 999


def setup_function():
    pass


def test_CT2_0x1015():
    assert set_ct2(conn_mode=modbus_config["conn_mode"], value=ct2) is False
    assert get_ct2(conn_mode=modbus_config["conn_mode"]) == 18


def teardown_function():
    pass
