import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

year = [1999, 2000, 2001, 2098, 2099, 2100]


def setup_function():
    global default_year
    default_year = int(get_real_time(conn_mode=modbus_config["conn_mode"])[0:4])
    logging.info('default year time ret is:{}'.format(default_year))


def test_year_0x1040():
    tcp_result = []
    for i in year:
        ret = set_year(conn_mode=modbus_config["conn_mode"], value=i)
        tcp_result.append(ret)
    assert tcp_result == [False, True, True, True, True, False]


def teardown_function():
    assert set_year(conn_mode=modbus_config["conn_mode"], value=default_year) is True
