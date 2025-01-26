import time

from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

slaveid = [1, 2, 246, 247, 248]


def setup_function():
    pass


def test_SlaveId_0x100E():
    rtu_slaveid = []
    for i in slaveid:
        time.sleep(1)
        ret = set_modbus_slaveid(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_slaveid.append(ret)
    logging.info('rtu_slaveid ret is:{}'.format(rtu_slaveid))
    for rtu in range(len(rtu_slaveid) - 1):
        assert rtu_slaveid[rtu] is True
    assert rtu_slaveid[-1] is False


def teardown_function():
    assert set_modbus_slaveid(conn_mode=modbus_config["conn_mode"], value=1)
