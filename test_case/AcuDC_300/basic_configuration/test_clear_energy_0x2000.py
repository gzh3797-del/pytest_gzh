from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

clear_energy_data = [0, 1, 2]


def setup_function():
    pass


def test_clear_energy_0x2000():
    assert clear_energy(conn_mode=modbus_config["conn_mode"], value=clear_energy_data[0]) is True
    assert clear_energy(conn_mode=modbus_config["conn_mode"], value=clear_energy_data[1]) is True
    assert clear_energy(conn_mode=modbus_config["conn_mode"], value=clear_energy_data[2]) is False


def teardown_function():
    pass
