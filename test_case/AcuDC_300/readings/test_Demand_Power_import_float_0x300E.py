from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Demand_Power_import_float_0x300E():
    assert set_demand_power_import_measurement(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_demand_power_import_measurement(conn_mode=modbus_config["conn_mode"])) is float


def teardown_function():
    pass
