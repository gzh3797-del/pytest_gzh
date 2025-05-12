from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

tcp_enable = [0, 1, 2, 0]


def setup_function():
    pass


def test_Tcp_Enable_0x1010():
    if modbus_config['conn_mode'] == 'rtu':
        tcp_tcp_enable = []
        for i in tcp_enable:
            ret = set_modbus_tcp_enable(conn_mode='rtu', value=i)
            tcp_tcp_enable.append(ret)
        logging.info('tcp_tcp_enable ret is:{}'.format(tcp_tcp_enable))
        for tcp in range(len(tcp_tcp_enable) - 1):
            assert tcp_tcp_enable[tcp] is True
        assert tcp_tcp_enable[-1] is False
    else:
        pass


def teardown_function():
    pass
