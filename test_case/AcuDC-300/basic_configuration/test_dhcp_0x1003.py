from comm.modbus_set_attr import *
from tools.log import Log
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

dhcp = [0, 1, 2]


def setup_function():
    pass


def test_dhcp_0x1003():
    rtu_dhcp = []
    for i in dhcp:
        ret = set_dhcp_enable(conn_mode=modbus_config["conn_mode"], value=i)
        rtu_dhcp.append(ret)
    assert rtu_dhcp[0] is True
    assert rtu_dhcp[1] is True
    assert rtu_dhcp[2] is False


def teardown_function():
    assert set_dhcp_enable(conn_mode=modbus_config["conn_mode"], value=0) is True
