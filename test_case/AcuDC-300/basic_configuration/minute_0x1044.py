import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

minute = [0, 1, 15, 58, 59, 60]


def setup_function():
    global default_minute
    default_minute = int(get_real_time(conn_mode=modbus_config["conn_mode"])[16:18])
    logging.info('default minute time ret is:{}'.format(default_minute))


def test_minute_0x1044():
    rtu_result = []
    for i in minute:
        ret = set_minute(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, True, False]


def teardown_function():
    assert set_minute(conn_mode=modbus_config["conn_mode"], value=default_minute) is True
