import time

from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

subnet_mask = ['0.0.0.0', '255.255.255.255', '256.256.256.256']


def setup_function():
    global default_dns_primary
    default_dns_primary = get_dns_primary_server(conn_mode=modbus_config["conn_mode"])
    logging.info('dns_primary_server is:{}'.format(default_dns_primary))


def test_DNS_primary_0x4106():
    rtu_dns_primary = []
    for i in subnet_mask:
        ret = set_dns_primary(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_dns_primary.append(ret)
    logging.info('rtu_dns_primary is:{}'.format(rtu_dns_primary))
    for index, element in enumerate(rtu_dns_primary):
        if index == 2:
            assert rtu_dns_primary[index] is False
        else:
            assert rtu_dns_primary[index] is True


def teardown_function():
    assert set_dns_primary(conn_mode=modbus_config["conn_mode"], value=default_dns_primary) is True
