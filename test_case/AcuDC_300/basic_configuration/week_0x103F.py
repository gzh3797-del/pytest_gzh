import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

week = [0, 1, 6, 8]


def setup_function():
    global default_week
    default_week = int(get_real_time(conn_mode=modbus_config["conn_mode"])[11:12])
    logging.info('default week time ret is:{}'.format(default_week))


def test_week_0x103F():
    rtu_result = []
    for i in week:
        ret = set_week(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, False]


def teardown_function():
    assert set_week(conn_mode=modbus_config["conn_mode"], value=default_week) is True
