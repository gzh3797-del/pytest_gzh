import logging
import time

from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

subnet_mask = ['0.0.0.0', '255.255.255.255', '256.256.256.256']


def setup_function():
    global default_gateway
    default_gateway = get_gateway(conn_mode=modbus_config["conn_mode"])
    logging.info('gateway is:{}'.format(default_gateway))


def test_gateway_0x4104():
    rtu_gateway = []
    for i in subnet_mask:
        ret = set_gateway(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_gateway.append(ret)
    logging.info('rtu_gateway is:{}'.format(rtu_gateway))
    for index, element in enumerate(rtu_gateway):
        if index == 2:
            assert rtu_gateway[index] is False
        else:
            assert rtu_gateway[index] is True


def teardown_function():
    assert set_gateway(conn_mode=modbus_config["conn_mode"], value=default_gateway) is True
