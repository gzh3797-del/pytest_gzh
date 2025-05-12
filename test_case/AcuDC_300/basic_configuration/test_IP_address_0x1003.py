from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

ip_address = ['193.168.1.254', '192.167.1.254', '192.168.2.254', '192.168.1.253', '193.169.3.250']


def setup_function():
    global default_ip_addr
    default_ip_addr = get_ip_address(conn_mode=modbus_config["conn_mode"])
    logging.info('default ip is:{}'.format(default_ip_addr))


def test_IP_address_0x1003():
    if modbus_config['conn_mode'] == 'rtu':
        rtu_ip = []
        for i in ip_address:
            ret = set_ip_address(conn_mode='rtu', value=i)
            rtu_ip.append(ret)
        for i in rtu_ip:
            assert i is True
    else:
        pass


def teardown_function():
    assert set_ip_address(conn_mode=modbus_config["conn_mode"], value=default_ip_addr) is True
