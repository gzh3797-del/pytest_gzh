from comm.modbus_set_attr import *
from comm.modbus_get_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])


def setup_function():
    pass


def test_Import_Charge_0x4010():
    assert set_import_charge_measurement(conn_mode=modbus_config["conn_mode"], value=0) is True
    assert type(read_import_charge_measurement(conn_mode=modbus_config["conn_mode"], ret={})) is float


def teardown_function():
    pass
