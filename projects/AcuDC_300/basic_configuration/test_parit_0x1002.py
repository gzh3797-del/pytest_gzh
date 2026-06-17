import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

parity = [0, 1, 2, 3, 4]


def setup_function():
    global default_parity
    default_parity = get_rs485_parity(conn_mode=modbus_config["conn_mode"])
    logging.info('default parity is:{}'.format(default_parity))


def test_parit_0x1002():
    rtu_parity = []
    for i in parity:
        time.sleep(1)
        ret = set_rs485_parity(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_parity.append(ret)
    for i in range(len(parity) - 1):
        assert rtu_parity[i] is True
    assert rtu_parity[-1] is False


def teardown_function():
    assert set_rs485_parity(conn_mode=modbus_config["conn_mode"], value=default_parity) is True
