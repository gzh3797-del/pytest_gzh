from comm.modbus_set_attr import *
from tools.log import Log

Log(str(__file__).split("\\")[-1])

meter_password = [0000, 9999, 10000]


def setup_function():
    pass


def test_password_0x1001():

    tcp_pass = []
    for password in meter_password:
        ret = set_meter_password(conn_mode=modbus_config["conn_mode"], value=password)
        tcp_pass.append(ret)
    for tcp in range(len(tcp_pass) - 1):
        assert tcp_pass[tcp] is True
    assert tcp_pass[-1] is False


def teardown_function():
    assert set_meter_password(conn_mode=modbus_config["conn_mode"], value=0000)
