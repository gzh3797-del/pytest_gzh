from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Demand_Current_export_int_0x3206():
    assert set_demand_current_export_client(conn_mode=modbus_config["conn_mode"], value=0) is False
    assert type(read_demand_current_export_client(conn_mode=modbus_config["conn_mode"])) is int


def teardown_function():
    pass
