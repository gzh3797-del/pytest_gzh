from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

rtu_enable = [0, 1, 2]


def setup_function():
    pass


def test_Rtu_Enable_0x100F():
    if modbus_config['conn_mode'] == 'tcp':
        tcp_rtu_enable = []
        for i in rtu_enable:
            ret = set_modbus_rtu_enable(conn_mode='tcp', value=i)
            tcp_rtu_enable.append(ret)
        logging.info('tcp_rtu_enable ret is:{}'.format(tcp_rtu_enable))
        for tcp in range(len(tcp_rtu_enable) - 1):
            assert tcp_rtu_enable[tcp] is True
        assert tcp_rtu_enable[-1] is False
    else:
        pass


def teardown_function():
    assert set_modbus_rtu_enable(conn_mode='tcp', value=1)
