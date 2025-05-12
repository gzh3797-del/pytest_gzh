from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_voltage_Measured_Compensated_int_0x3200():
    assert set_voltage_measu_or_comp_client(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_voltage_measurement_client(conn_mode=modbus_config["conn_mode"])) is int


def teardown_function():
    pass
