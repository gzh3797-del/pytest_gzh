from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

ct1 = 999


def setup_function():
    pass


def test_CT1_0x1014():
    assert set_ct1(conn_mode=modbus_config["conn_mode"], value=ct1) is False
    assert get_ct1(conn_mode=modbus_config["conn_mode"]) == 650


def teardown_function():
    pass
