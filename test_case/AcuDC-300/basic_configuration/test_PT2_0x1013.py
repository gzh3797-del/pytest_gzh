from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

pt2 = 999


def setup_function():
    pass


def test_PT2_0x1013():
    assert set_pt2(conn_mode=modbus_config["conn_mode"], value=pt2) is False
    assert get_pt2(conn_mode=modbus_config["conn_mode"]) == 1000


def teardown_function():
    pass
