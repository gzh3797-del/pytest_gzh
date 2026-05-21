import time

from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Voltage_Compensated_float_0x3014():
    assert set_voltage_compensated(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_voltage_compensated(conn_mode=modbus_config["conn_mode"])) is float


def teardown_function():
    pass
