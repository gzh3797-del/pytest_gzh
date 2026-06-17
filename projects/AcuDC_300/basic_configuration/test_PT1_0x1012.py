from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

pt1 = 999


def setup_function():
    pass


def test_PT1_0x1012():
    assert set_pt1(conn_mode=modbus_config["conn_mode"], value=pt1) is False
    assert get_pt1(conn_mode=modbus_config["conn_mode"]) == 1000


def teardown_function():
    pass
