import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

backlight_time = [0, 1, 3, 119, 120, 121]


def setup_function():
    global default_backlight_time
    default_backlight_time = get_backlight_time(conn_mode=modbus_config["conn_mode"])
    logging.info('default backlight time ret is:{}'.format(backlight_time))


def test_backlight_time_0x101C():
    tcp_result = []
    for i in backlight_time:
        ret = set_backlight_time(conn_mode=modbus_config["conn_mode"], value=i)
        tcp_result.append(ret)
    assert tcp_result == [True, True, True, True, True, False]


def teardown_function():
    assert set_backlight_time(conn_mode=modbus_config["conn_mode"], value=default_backlight_time) is True
