from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Total_Charge_0x401C():
    assert set_total_charge_measurement(conn_mode=modbus_config["conn_mode"], value=0) is True
    assert type(read_total_charge_measurement(conn_mode=modbus_config["conn_mode"], ret={})) is float


def teardown_function():
    pass
