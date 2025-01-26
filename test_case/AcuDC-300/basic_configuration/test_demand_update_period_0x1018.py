from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

demand_update_period = [0, 1, 2, 29, 30, 31]


def setup_function():
    global default_demand_update_period
    default_demand_update_period = get_demand_update_period(conn_mode=modbus_config["conn_mode"])
    logging.info('default demand update period ret is:{}'.format(demand_update_period))


def test_demand_update_period_0x1018():
    rtu_result = []
    for i in demand_update_period:
        ret = set_demand_update_period(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, True, True, False]


def teardown_function():
    assert set_demand_update_period(conn_mode=modbus_config["conn_mode"], value=default_demand_update_period) is True
