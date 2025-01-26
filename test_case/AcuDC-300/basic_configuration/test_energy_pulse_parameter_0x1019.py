from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

energy_pulse_parameter = [0, 1, 2, 3, 4, 5]


def setup_function():
    global default_energy_pulse_parameter
    default_energy_pulse_parameter = get_energy_pulse_para(conn_mode=modbus_config["conn_mode"])
    logging.info('default energy pulse constant ret is:{}'.format(energy_pulse_parameter))


def test_energy_pulse_parameter_0x1019():
    rtu_result = []
    for i in energy_pulse_parameter:
        ret = set_energy_pulse_parameter(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, True, False]


def teardown_function():
    assert set_energy_pulse_parameter(conn_mode=modbus_config["conn_mode"], value=default_energy_pulse_parameter) is True
