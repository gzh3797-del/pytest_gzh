from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

clear_device_run_time_data = [0, 1, 2]


def setup_function():
    pass


def test_clear_device_run_time_0x2004():
    assert clear_device_run_time(conn_mode=modbus_config["conn_mode"], value=clear_device_run_time_data[0]) is True
    assert clear_device_run_time(conn_mode=modbus_config["conn_mode"], value=clear_device_run_time_data[1]) is True
    assert clear_device_run_time(conn_mode=modbus_config["conn_mode"], value=clear_device_run_time_data[2]) is False


def teardown_function():
    pass
