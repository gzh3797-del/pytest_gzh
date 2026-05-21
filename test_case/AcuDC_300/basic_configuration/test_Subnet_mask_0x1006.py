import logging

from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

subnet_mask = ['0.0.0.0', '254.254.254.254', '256.256.256.256']


def setup_function():
    global default_subnet_mask
    default_subnet_mask = get_subnet_mask(conn_mode=modbus_config["conn_mode"])
    logging.info('subnet mask is:{}'.format(default_subnet_mask))


def test_Subnet_mask_0x1006():
    rtu_ip = []
    for i in subnet_mask:
        ret = set_subnet_mask(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_ip.append(ret)
    logging.info('rtu_ip is:{}'.format(rtu_ip))
    for index, element in enumerate(rtu_ip):
        if index == 2:
            assert rtu_ip[index] is False
        else:
            assert rtu_ip[index] is True


def teardown_function():
    assert set_subnet_mask(conn_mode=modbus_config["conn_mode"], value=default_subnet_mask) is True
