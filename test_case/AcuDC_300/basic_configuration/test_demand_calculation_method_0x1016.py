from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

demand_calculation_method = [0, 1, 2]


def setup_function():
    global default_demand_calculation_method
    default_demand_calculation_method = get_demand_calculation_method(conn_mode=modbus_config["conn_mode"])
    logging.info('default demand calculation method ret is:{}'.format(default_demand_calculation_method))


def test_demand_calculation_method_0x1016():
    rtu_result = []
    for i in demand_calculation_method:
        ret = set_demand_calculation_method(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result[0] is True
    assert rtu_result[1] is True
    assert rtu_result[2] is False


def teardown_function():
    assert set_demand_calculation_method(conn_mode=modbus_config["conn_mode"], value=default_demand_calculation_method) is True
