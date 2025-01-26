import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

second = [0, 1, 15, 58, 59, 60]


def setup_function():
    global default_second
    default_second = int(get_real_time(conn_mode=modbus_config["conn_mode"])[19:21])
    logging.info('default second time ret is:{}'.format(default_second))


def test_second_0x1045():
    rtu_result = []
    for i in second:
        time.sleep(1)
        ret = set_second(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_result.append(ret)
    assert rtu_result == [True, True, True, True, True, False]


def teardown_function():
    time.sleep(2)
    assert set_second(conn_mode=modbus_config["conn_mode"], value=default_second) is True
