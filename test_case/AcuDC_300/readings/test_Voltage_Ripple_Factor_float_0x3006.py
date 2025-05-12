from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Voltage_Ripple_Factor_float_0x3006():
    assert set_voltage_ripple_factor_measurement(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_voltage_ripple_factor_measurement(conn_mode=modbus_config["conn_mode"])) is float


def teardown_function():
    pass
