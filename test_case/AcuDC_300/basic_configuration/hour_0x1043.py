import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

hour = [0, 1, 15, 22, 23, 24]


def setup_function():
    global default_hour
    default_hour = int(get_real_time(conn_mode=modbus_config["conn_mode"])[13:15])
    logging.info('default hour time ret is:{}'.format(default_hour))


def test_hour_0x1043():
    rtu_result = []
    for i in hour:
        ret = set_hour(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, True, False]


def teardown_function():
    assert set_hour(conn_mode=modbus_config["conn_mode"], value=default_hour) is True
