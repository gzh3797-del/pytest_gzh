from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

seal_status = 0


def setup_function():
    pass


def test_seal_status_0x101D():
    assert set_seal_status(conn_mode=modbus_config["conn_mode"], value=seal_status) is False
    assert get_seal_status(conn_mode=modbus_config["conn_mode"]) == 0 or get_seal_status(conn_mode='tcp') == 1


def teardown_function():
    pass
