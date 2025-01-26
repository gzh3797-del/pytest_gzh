from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

energy_pulse_constant = [0.09, 0.1, 0.2, 99999, 100000, 100000.1]


def setup_function():
    global default_energy_pulse_constant
    default_energy_pulse_constant = get_energy_pulse_constant(conn_mode=modbus_config["conn_mode"])
    logging.info('default energy pulse constant ret is:{}'.format(energy_pulse_constant))


def test_energy_pulse_constant_0x101A():
    rtu_result = []
    for i in energy_pulse_constant:
        ret = set_energy_pulse_constant(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, True, True, False]


def teardown_function():
    assert set_energy_pulse_constant(conn_mode=modbus_config["conn_mode"], value=default_energy_pulse_constant) is True
