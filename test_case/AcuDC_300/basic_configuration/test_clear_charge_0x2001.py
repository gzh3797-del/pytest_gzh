from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

clear_charge_data = [0, 1, 2]


def setup_function():
    pass


def test_clear_charge_0x2001():
    assert clear_charge(conn_mode=modbus_config["conn_mode"], value=clear_charge_data[0]) is True
    assert clear_charge(conn_mode=modbus_config["conn_mode"], value=clear_charge_data[1]) is True
    assert clear_charge(conn_mode=modbus_config["conn_mode"], value=clear_charge_data[2]) is False


def teardown_function():
    pass
