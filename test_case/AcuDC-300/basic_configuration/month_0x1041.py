import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

month = [0, 1, 2, 11, 12, 13]


def setup_function():
    global default_month
    default_month = int(get_real_time(conn_mode=modbus_config["conn_mode"])[5:7])
    logging.info('default month time ret is:{}'.format(default_month))


def test_month_0x1041():
    rtu_result = []
    for i in month:
        ret = set_month(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [False, True, True, True, True, False]


def teardown_function():
    time.sleep(1)
    assert set_month(conn_mode=modbus_config["conn_mode"], value=default_month) is True
