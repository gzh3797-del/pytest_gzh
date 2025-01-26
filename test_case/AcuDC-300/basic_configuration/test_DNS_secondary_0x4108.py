from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

subnet_mask = ['0.0.0.0', '255.255.255.255', '256.256.256.256']


def setup_function():
    global default_dns_secondary
    default_dns_secondary = get_dns_secondary_server(conn_mode=modbus_config["conn_mode"])
    logging.info('default_dns_secondary is:{}'.format(default_dns_secondary))


def test_DNS_secondary_0x4108():
    rtu_dns_secondary = []
    for i in subnet_mask:
        ret = set_dns_secondary(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_dns_secondary.append(ret)
    logging.info('rtu_dns_secondary is:{}'.format(rtu_dns_secondary))
    for index, element in enumerate(rtu_dns_secondary):
        if index == 2:
            assert rtu_dns_secondary[index] is False
        else:
            assert rtu_dns_secondary[index] is True


def teardown_function():
    assert set_dns_secondary(conn_mode=modbus_config["conn_mode"], value=default_dns_secondary) is True
