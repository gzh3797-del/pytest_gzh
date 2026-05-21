import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

day = [0, 1, 2, 30, 31, 32]


def setup_function():
    global default_day
    default_day = int(get_real_time(conn_mode=modbus_config["conn_mode"])[8:10])
    logging.info('default day time ret is:{}'.format(default_day))


def test_day_0x1042():
    rtu_result = []
    for i in day:
        ret = set_day(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, True, True, False]


def teardown_function():
    time.sleep(1)
    assert set_day(conn_mode='tcp', value=default_day) is True
