from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Demand_Current_export_float_0x300C():
    assert set_demand_current_export_measurement(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_demand_current_export_measurement(conn_mode=modbus_config["conn_mode"])) is float


def teardown_function():
    pass
