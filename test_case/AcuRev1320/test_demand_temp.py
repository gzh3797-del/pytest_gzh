import pytest

from comm.AcuRev4100_modbus_get_attr import *
from comm.source_control import *

import math

from acurev1320_modbus_get import *
from test_case.AcuRev1320.fast_test.memory_addrs import MemoryAddr


# 全局 ModbusClient，只初始化一次

@pytest.fixture(scope="function")  # 每个test function,都独立准备环境，例如每次建立Modbus/TCP连接
def modbus_client():
    """
    每个测试函数执行前：
    - 切换交流界面
    - 档位归零
    - 等待稳定
    测试结束后：
    - 关源
    - 切回默认界面
    """
    ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])
    switch_device_screen_interface(inter=0x01)# 切换至交流界面
    set_gear_switching_mode('00000000')# 档位切换归零
    time.sleep(5)

    yield  ModbusClient # 测试执行

    up_source_ac()# 关源
    switch_device_screen_interface(inter=0x00)# 切换至默认界面
    ModbusClient.close()

# ------------------- 操作函数 -------------------
def set_wire_type(ModbusClient,voltage_wire_value=4):
    """
    设置接线方式
    :param voltage_wire_value: 电压接线方式
        0: ELEMENT_1_WIRE_2
        1: ELEMENT_2_WIRE_3_PHASE_1
        2: ELEMENT_2_WIRE_3_DELTA
        3: ELEMENT_2_WIRE_3_NETWORK
        4: ELEMENT_3_WIRE_4_Y
        5：ELEMENT_3_WIRE_4_DELTA
    :return:
    """
    ret = ModbusClient.write_registers(addr=MemoryAddr.voltage_wire_addr, value=voltage_wire_value)

def set_para_modbus(ModbusClient,addr, value):
    """
    寄存器配置-电压接线方式
    :param addr: 寄存器地址
    :param value: 值
    :return: True:写入成功,False:写入失败
    """
    ret = ModbusClient.write_registers(address=addr, values=value, slave=1)
    if f'{(addr, 1)}' not in str(ret):
        logging.error('Set_Service_Configuration fail, ret is:{}'.format(ret))
        return False
    ret = ModbusClient.read_measurement(address=addr, count=1, slave=1)
    if ret[0] == value:
        return True
    return False


def set_demand_para(ModbusClient,demand_method, demand_interval, demand_update_rate):
    """

    Args:
        demand_method: Fixed Window: 0  Sliding Window: 1
        demand_interval: 1~30 minute
        demand_update_rate: 1~30 minute
    Returns:

    """
    # 设置demand method, sliding or fixed
    set_para_modbus(ModbusClient,MemoryAddr.demand_update_rate_addr, demand_method)
    # 设置demand interval(min)
    set_para_modbus(ModbusClient, MemoryAddr.demand_interval_addr, demand_interval)
    # 设置demand update rate(min)
    set_para_modbus(ModbusClient, MemoryAddr.demand_update_rate_addr, demand_update_rate)


def voltage_current(angle=0, current=2):
    """
    Args:
        angle: 电压电流夹角（单位：度）
        current: 电流值
    Returns:
        None
    """
    set_ac(
        120 + angle, 240 + angle, 0 + angle,
        120, 240, 0,
        400, 400, 400,
        current, current, current,
        50
    )

def read_demand(ModbusClient,standard_demand_value, demand_address=0xC466, tolerance=0.01):
    """
    Args:
        standard_value (float): 预期需量值
        demand_address (int): 需量寄存器地址
        tolerance (float): 允许相对误差，默认 1% (=0.01)
    Returns:
        value_measu (float): 实际测量值
        is_pass (bool): 误差是否在允许范围内
        error_percent (float): 相对误差百分比
    """
    value = ModbusClient.read_measurement(demand_address, reg_count=2, slave=1)
    logging.info('The value of register address %s is: %s', hex(demand_address), value)
    # 2 个 16bit 寄存器拼成 32bit
    reg_hex = (
            hex(value[0]).replace('0x', '').zfill(4) +
            hex(value[1]).replace('0x', '').zfill(4)
    )
    integer_num = int(reg_hex, 16)
    # 按 IEEE754 float 解析
    value_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
    # ---------- 误差计算 ----------
    if standard_demand_value != 0:
        error_percent = abs(value_measu - standard_demand_value) / standard_demand_value * 100
        is_pass = error_percent <= (tolerance * 100)
        return value_measu, is_pass, error_percent
    if standard_demand_value == 0 and abs(value_measu) < 0.01:  # 用于判断demand清除是否成功
        return value_measu, True, 0


def check_demand_clear(ModbusClient):
    """
    0xC466 System Active Power Demand
    0xC468 System Reactive Power Demand
    0xC46A System Apparent Power Demand
    0xC46C Phase A Current Demand
    0xC46E Phase B Current Demand
    0xC470 Phase C Current Demand
    0xC472 Phase N Current Demand
    0xC474 Predictive System Active Power Demand
    0xC476 Predictive System Reactive Power Demand
    0xC478 Predictive System Apparent Power Demand
    Returns:

    """
    value = 0
    for name, addr in MemoryAddr.demand_addr.items():
        value_measu, _, _ = read_demand(ModbusClient,0, addr, 0.01)
        if abs(value_measu) >= 0.001:
            return False
    return True

def demand_diff_read(ModbusClient,current,angle):
    rad = math.radians(angle)  # 转为弧度
    re1 = read_demand(ModbusClient,400*current*math.cos(rad),MemoryAddr.demand_addr.system_active_power,0.1)[1]
    re2 = read_demand(ModbusClient,400*current*math.sin(rad), MemoryAddr.demand_addr.system_reactive_power, 0.1)[1]
    re3 = read_demand(ModbusClient,400*current*3, MemoryAddr.demand_addr.system_apparent_power, 0.1)[1]
    re4 = read_demand(ModbusClient,current, MemoryAddr.demand_addr.phase_a_current, 0.1)[1]
    re5 = read_demand(ModbusClient,current, MemoryAddr.demand_addr.phase_b_current, 0.1)[1]
    re6 = read_demand(ModbusClient,current, MemoryAddr.demand_addr.phase_c_current, 0.1)[1]
    re7 = read_demand(ModbusClient,0, MemoryAddr.demand_addr.phase_n_current, 0.1)[1]
    return all([re1, re2, re3, re4, re5, re6, re7])

# ------------------- 测试函数 -------------------
def test_demand_case(ModbusClient, angle_value, current_value, wire_type, demand_method, demand_interval, demand_update_rate):
    # 设置接线方式
    set_wire_type(ModbusClient, wire_type)
    # 设置电压、电流
    voltage_current(angle_value, current_value)
    # 设置需量参数
    set_demand_para(ModbusClient, demand_method, demand_interval, demand_update_rate)

    # 检查需量是否清零
    assert check_demand_clear(ModbusClient), "需量未清零"

    # 第一个周期不计算
    time.sleep(demand_interval * 60)
    # 第二个周期开始计算
    time.sleep(demand_interval * 60)
    assert demand_diff_read(ModbusClient, current_value, angle_value), "第一个周期需量计算错误"

    # 0.5*demand_interval后降低电流
    time.sleep(0.5 * demand_interval * 60)
    voltage_current(angle_value, 0.5 * current_value)
    # 检查需量是否正确
    assert demand_diff_read(ModbusClient, 0.75 * current_value, angle_value), "第二个周期需量计算错误"


# ------------------- 动态生成测试 -------------------
class TestDemand:
    @classmethod
    def generate_tests_fixed(cls):
        pos_dic = [[30, 4, 4, 0, 4, 10],
                   [30, 4, 4, 0, 4, 10],
                   [30, 4, 4, 0, 4, 10],
                   [30, 4, 4, 0, 4, 10],
                   [30, 4, 4, 0, 4, 10]]

        for i, params in enumerate(pos_dic):
            angle_value, current_value, wire_type, demand_method, demand_interval, demand_update_rate = params

            def make_test(angle_value, current_value, wire_type, demand_method, demand_interval, demand_update_rate):
                def test(modbus_client):  # 保留 fixture 参数
                    return test_demand_case(modbus_client, angle_value, current_value, wire_type,
                                            demand_method, demand_interval, demand_update_rate)
                return test

            test_name = f"test_{angle_value}deg_{current_value}A_wire{wire_type}_Fixed_{demand_interval}min_updaterate_{demand_update_rate}min_{i}"
            setattr(cls, test_name, make_test(angle_value, current_value, wire_type, demand_method, demand_interval, demand_update_rate))

TestDemand.generate_tests_fixed()

# ------------------- 脚本执行 -------------------
if __name__ == "__main__":
        pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html", '-x'])








