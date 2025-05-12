from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

demand_window_time = [0, 1, 2, 29, 30, 31]


def setup_function():
    global default_demand_window_time
    default_demand_window_time = get_demand_window_time(conn_mode=modbus_config["conn_mode"])
    logging.info('default demand window time ret is:{}'.format(default_demand_window_time))


def test_demand_window_time_0x1017():
    rtu_result = []
    for i in demand_window_time:
        ret = set_demand_window_time(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, True, True, False]


def teardown_function():
    assert set_demand_window_time(conn_mode=modbus_config["conn_mode"], value=default_demand_window_time) is True
