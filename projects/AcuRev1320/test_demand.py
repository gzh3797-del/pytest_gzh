import pytest
from datetime import datetime
import struct
from comm.source_control import *
import math

from projects.AcuRev1320.fast_test.acuvimseries_modbus_get import  *
from projects.AcuRev1320.fast_test.memory_addrs import MemoryAddr

# 全局 ModbusClient，只初始化一次
ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


@pytest.fixture(scope="function", autouse=True)  # 每个test function,都独立准备环境，例如每次建立Modbus/TCP连接
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
    switch_device_screen_interface(inter=0x01)  # 切换至交流界面
    set_gear_switching_mode('00000000')  # 档位切换归零
    time.sleep(5)

    yield  # 测试执行

    up_source_ac()  # 关源
    switch_device_screen_interface(inter=0x00)  # 切换至默认界面


# ------------------- 操作函数 -------------------
def set_wire_type(voltage_wire_value=4):
    """
    设置接线方式
    :param voltage_wire_value: 电压接线方式
        0: ELEMENT_1_WIRE_2
        1: ELEMENT_2_WIRE_3_PHASE_1
        2: ELEMENT_2_WIRE_3_DELTA
        3: ELEMENT_2_WIRE_3_NETWORK
        4: ELEMENT_3_WIRE_4_Y
        5: ELEMENT_3_WIRE_4_DELTA
    :return:
    """
    ModbusClient.write_registers(MemoryAddr.voltage_wire_addr, voltage_wire_value)


def set_para_modbus(addr, value):
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


def set_demand_para(demand_method, demand_interval, demand_update_rate):
    """

    Args:
        demand_method: Fixed Window: 0  Sliding Window: 1
        demand_interval: 1~30 minute
        demand_update_rate: 1~30 minute
    Returns:

    """
    # 设置demand method, sliding or fixed
    set_para_modbus(MemoryAddr.demand_algorithm_addr, demand_method)
    # 设置demand interval(min)
    set_para_modbus(MemoryAddr.demand_interval_addr, demand_interval)
    # 设置demand update rate(min)
    set_para_modbus(MemoryAddr.demand_update_rate_addr, demand_update_rate)


def set_time_trigger():
    # 设置系统时间毫秒
    set_para_modbus(MemoryAddr.sys_millisecond, 1)


def reset_demand_trigger():
    # 重置系统最大需量值，重置需量
    set_para_modbus(MemoryAddr.clear_max_demand, 1)


def voltage_current(voltage=0, angle=0, current=2):
    """
    Args:
        voltage: 设置电压值（单位：V）
        angle: 电压电流夹角（单位：度）
        current: 设置电压值（单位：A）
    Returns:
        None
    """
    set_ac(
        120 + angle, 240 + angle, 0 + angle,
        120, 240, 0,
        voltage, voltage, voltage,
        current, current, current,
        50
    )


def read_demand(standard_demand_value, demand_address=0xC466, tolerance=0.01):
    """
    Args:
        standard_demand_value (float): 预期需量值
        demand_address (int): 需量寄存器地址
        tolerance (float): 允许相对误差，默认 1% (=0.01)
    Returns:
        value_measu (float): 实际测量值
        is_pass (bool): 误差是否在允许范围内
        error_percent (float): 相对误差百分比
    """
    value = ModbusClient.read_measurement(demand_address, 2, 1)
    logging.info('The value of register address %s is: %s', hex(demand_address), value)

    # 2 个 16bit 寄存器拼成 32bit
    reg_hex = (
            hex(value[0]).replace('0x', '').zfill(4) +
            hex(value[1]).replace('0x', '').zfill(4)
    )
    integer_num = int(reg_hex, 16)
    # 按 IEEE754 float 解析
    value_measure = struct.unpack('!f', struct.pack('!I', integer_num))[0]

    # ---------- 误差计算 ----------
    if standard_demand_value != 0:
        error_percent = abs(value_measure - standard_demand_value) / standard_demand_value * 100
        is_pass = error_percent <= (tolerance * 100)
        return value_measure, is_pass, error_percent
    # 用于判断demand清除是否成功
    if standard_demand_value == 0:
        is_pass = abs(value_measure) < 0.01
        return value_measure, is_pass, 0 if is_pass else value_measure


def check_demand_clear():
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
    for name, addr in MemoryAddr.demand_addr.items():
        value_measure, _, _ = read_demand(0, addr, 0.01)
        if abs(value_measure) >= 0.001:
            return False
    return True


def demand_cal_result_get(predict_system_demand_p, predict_system_demand_q, predict_system_demand_s, demand_current):
    """

    Args:
        predict_system_demand_p: 预期系统有功需量
        predict_system_demand_q: 预期系统无功需量
        predict_system_demand_s: 预期系统视在需量
        demand_current: 预期电流需量

    Returns:

    """
    re1 = read_demand(predict_system_demand_p, MemoryAddr.demand_addr.system_active_power,0.1)[1]
    re2 = read_demand(predict_system_demand_q, MemoryAddr.demand_addr.system_reactive_power, 0.1)[1]
    re3 = read_demand(predict_system_demand_s, MemoryAddr.demand_addr.system_apparent_power, 0.1)[1]
    re4 = read_demand(demand_current, MemoryAddr.demand_addr.phase_a_current, 0.1)[1]
    re5 = read_demand(demand_current, MemoryAddr.demand_addr.phase_b_current, 0.1)[1]
    re6 = read_demand(demand_current, MemoryAddr.demand_addr.phase_c_current, 0.1)[1]
    re7 = read_demand(demand_current, MemoryAddr.demand_addr.phase_n_current, 0.1)[1]
    return all([re1, re2, re3, re4, re5, re6, re7])


def calc_wait_seconds(demand_interval_min: int) -> int:
    """

    Args:
        demand_interval_min: 需量窗口间隔时间

    Returns: 当前时刻基于demand_interval_min等待时间，s

    """
    now = datetime.now()
    print(f"当前时间:{now}")
    seconds_today = now.hour * 3600 + now.minute * 60 + now.second
    interval_seconds = demand_interval_min * 60
    next_boundary = math.ceil(seconds_today / interval_seconds) * interval_seconds
    wait_seconds = next_boundary - seconds_today
    return wait_seconds if wait_seconds != 0 else interval_seconds


# ---------------------测试函数------------------------


def generate_demand_fixed(voltage, angle_value, current_value, wire_type,
                          demand_method, demand_interval, demand_update_rate, demand_trigger):
    """

    Args:
        voltage:电压值
        angle_value:压流角
        current_value:电流值
        wire_type: 0: ELEMENT_1_WIRE_2
                   1: ELEMENT_2_WIRE_3_PHASE_1
                   2: ELEMENT_2_WIRE_3_DELTA
                   3: ELEMENT_2_WIRE_3_NETWORK
                   4: ELEMENT_3_WIRE_4_Y
                   5：ELEMENT_3_WIRE_4_DELTA
        demand_method: 需量计算算法
        demand_interval: 需量计算间隔
        demand_update_rate: 需量上报时间间隔
        demand_trigger:0表示 修改参数触发需量重置，1表示修改时间触发需量重置，2表示重置系统最大需量值来重置需量

    Returns:

    """
    # 初始需量值
    rad = math.radians(angle_value)  # 转为弧度
    predict_system_demand_p = voltage * current_value * math.cos(rad) * 3
    predict_system_demand_q = voltage * current_value * math.sin(rad) * 3
    predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
    demand_current = current_value

    # 初始需量测量需量设置
    set_wire_type(wire_type)  # 设置接线方式
    voltage_current(voltage, angle_value, current_value)  # 设置电压、电流
    set_demand_para(demand_method, demand_interval, demand_update_rate)  # 设置需量参数
    time.sleep(300)
    assert not check_demand_clear(), "需量预制数据不为0"

    # 触发需量重置
    if demand_trigger == 0:
        set_demand_para(demand_method, demand_interval, demand_update_rate)  # 设置需量参数
    elif demand_trigger == 1:    # 设置时间重置需量
        set_time_trigger()
    else:
        reset_demand_trigger()  # 设置重置需量最大值来重置需量
    # 检查需量值是否为0
    assert check_demand_clear(), "设置需量参数后，需量没有置为0"

    wait_time = calc_wait_seconds(demand_interval)  # 计算基于当前时间最近的demand_interval整数倍的时刻，需等待时间(s)
    print(f"等待: {wait_time}s")
    time.sleep(wait_time)  # 距离最近的demand_interval整数倍时刻期间的时间，不计算需量。
    time.sleep(demand_interval*60)  # 第二个周期开始计算
    if voltage < 9.5 or current_value == 0:
        assert check_demand_clear(), "设置需量参数后，需量没有置为0"
    else:
        assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                     predict_system_demand_s, demand_current), "第一个周期需量计算错误"
        # 0.5*demand_interval后，降低电流到0.5*current_value
        time.sleep(0.5*demand_interval * 60)
        voltage_current(voltage*0.5, angle_value, 0.5*current_value)
        # 0.5*demand_interval后，修改电压为0.5*voltage，修改电流为0.5*I，
        # 查看需量是否正确，voltage*0.625用于计算功率需量，current_value*0.75用于计算电流需量
        predict_system_demand_p = voltage * current_value * 0.625 * math.cos(rad) * 3
        predict_system_demand_q = voltage * current_value * 0.625 * math.sin(rad) * 3
        predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
        demand_current = current_value * 0.5
        assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                     predict_system_demand_s, demand_current), "第二个周期需量计算错误"


def generate_demand_sliding(voltage, angle_value, current_value, wire_type,
                            demand_method, demand_interval, demand_update_rate, demand_trigger):
    """

    Args:
        voltage:电压值
        angle_value:压流角
        current_value:电流值
        wire_type: 0: ELEMENT_1_WIRE_2
                   1: ELEMENT_2_WIRE_3_PHASE_1
                   2: ELEMENT_2_WIRE_3_DELTA
                   3: ELEMENT_2_WIRE_3_NETWORK
                   4: ELEMENT_3_WIRE_4_Y
                   5：ELEMENT_3_WIRE_4_DELTA
        demand_method: 需量计算算法
        demand_interval: 需量计算间隔
        demand_update_rate: 需量上报时间间隔
        demand_trigger:0表示 修改参数触发需量重置，1表示修改时间触发需量重置

    Returns:

    """
    # 初始需量值
    rad = math.radians(angle_value)  # 转为弧度
    predict_system_demand_p = voltage * current_value * math.cos(rad) * 3
    predict_system_demand_q = voltage * current_value * math.sin(rad) * 3
    predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
    demand_current = current_value

    # 初始需量测量需量设置
    set_wire_type(wire_type)  # 设置接线方式
    voltage_current(voltage, angle_value, current_value)  # 设置电压、电流
    set_demand_para(0, 2, 2)  # 设置需量参数，产生不为0需量值
    time.sleep(300)
    assert not check_demand_clear(), "需量预制数据不为0"

    # 触发需量重置
    if demand_trigger == 0:
        set_demand_para(demand_method, demand_interval, demand_update_rate)  # 设置需量参数，触发需量清零
    elif demand_trigger == 1:
        set_time_trigger()  # 设置时间重置需量
    else:
        reset_demand_trigger()  # 设置重置需量最大值来重置需量
    assert check_demand_clear(), "设置需量参数后，需量没有置为0"

    wait_time = calc_wait_seconds(demand_interval)  # 计算基于当前时间最近的demand_interval整数倍的时刻，需等待时间(s)
    print(f"等待: {wait_time}s")
    time.sleep(wait_time)  # 距离最近的demand_interval整数倍时刻期间的时间，不计算需量。
    time.sleep(demand_interval*60)  # 第二个周期开始计算

    if voltage < 9.5 or current_value == 0 or (current_value == 0 & voltage < 9.5):
        assert check_demand_clear(), "设置需量参数后，需量没有置为0"
    else:
        # 获取比对第二个周期需量值
        assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                     predict_system_demand_s, demand_current), "第一个周期需量计算错误"

        # 0.5*demand_interval后，降低电流到0.5*current_value
        voltage_current(voltage * 0.5, angle_value, 0.5 * current_value)
        time.sleep(demand_update_rate * 60)
        if demand_update_rate < demand_interval:
            # 获取比对第二个周期需量值
            predict_system_demand_p = voltage * current_value * math.cos(rad)*3*(1-0.75*(demand_update_rate/demand_interval))
            predict_system_demand_q = voltage * current_value * math.sin(rad)*3*(1-0.75*(demand_update_rate/demand_interval))
            predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
            demand_current = current_value * 0.5
            assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                         predict_system_demand_s, demand_current), "第二个周期需量计算错误"
            # 获取比对第三个周期需量值
            time.sleep(demand_update_rate * 60)
            if demand_update_rate*2 >= demand_interval:
                predict_system_demand_p = voltage * 0.5 * current_value * 0.5 * math.cos(rad) * 3
                predict_system_demand_q = voltage * 0.5 * current_value * 0.5 * math.sin(rad) * 3
                predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
                demand_current = current_value * 0.5
                assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                             predict_system_demand_s, demand_current), "第三个周期需量计算错误"
            else:
                predict_system_demand_p = voltage * current_value * math.cos(rad) * 3 * (1 - 1.5 * (demand_update_rate / demand_interval))
                predict_system_demand_q = voltage * current_value * math.sin(rad) * 3 * (1 - 1.5 * (demand_update_rate / demand_interval))
                predict_system_demand_s = math.sqrt(math.pow(predict_system_demand_p/3, 2)+math.pow(predict_system_demand_q/3, 2)) * 3
                demand_current = current_value * (1-(demand_update_rate / demand_interval))
                assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                             predict_system_demand_s, demand_current), "第三个周期需量计算错误"
        else:
            # 获取比对第二个周期需量值
            predict_system_demand_p = voltage * 0.5 * current_value * 0.5 * math.cos(rad) * 3
            predict_system_demand_q = voltage * 0.5 * current_value * 0.5 * math.sin(rad) * 3
            predict_system_demand_s = math.sqrt(
                math.pow(predict_system_demand_p / 3, 2) + math.pow(predict_system_demand_q / 3, 2)) * 3
            demand_current = current_value * 0.5
            assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                         predict_system_demand_s, demand_current), "第二个周期需量计算错误"

            # 获取比对第三个周期需量值
            time.sleep(demand_update_rate*60)
            assert demand_cal_result_get(predict_system_demand_p, predict_system_demand_q,
                                         predict_system_demand_s, demand_current), "第三个周期需量计算错误"

# ---------------------测试数据------------------------


input_fixed = [
    [200, 0, 2, 4, 0, 4, 15, 0],
    [9, 0, 2, 4, 0, 5, 15, 0],
    [200, 0, 2, 4, 0, 10, 10, 1],
    [200, 30, 2, 4, 0, 5, 10, 1],
    [200, 120, 2, 4, 0, 15, 15, 2],
    [200, 120, 2, 4, 0, 30, 15, 2],
    [200, 210, 2, 4, 0, 5, 15, 1],
    [200, 330, 2, 4, 0, 7, 15, 0],
    [200, 0, 0, 4, 0, 5, 15, 0],
    [0, 0, 2, 4, 0, 5, 15, 0]
]

input_sliding = [
    [200, 0, 2, 4, 1, 10, 5, 0],
    [200, 30, 2, 4, 1, 15, 5, 0],
    [200, 120, 2, 4, 1, 5, 5, 1],
    [200, 210, 2, 4, 1, 6, 10, 1],
    [200, 330, 2, 4, 1, 30, 15, 2],
    [200, 0, 0, 4, 1, 5, 5, 0],
    [9, 0, 2, 4, 1, 5, 5, 0]
]


# --------------------- 生成 IDs ---------------------
def make_ids(data_list, mode_name):
    """根据数组生成可读 pytest ID"""
    ids = []
    for row in data_list:
        voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger = row
        ids.append(f"{mode_name}_V{voltage}_A{angle}_interval{interval}_update{update_rate}")
    return ids


fixed_ids = make_ids(input_fixed, "fixed")
sliding_ids = make_ids(input_sliding, "sliding")


# --------------------- pytest 参数化 ---------------------
@pytest.mark.parametrize(
    "voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger",
    input_fixed,
    ids=fixed_ids
)
def test_demand_fixed(voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger):
    # 调用你的 generate_demand_fixed 函数
    generate_demand_fixed(voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger)


@pytest.mark.parametrize(
    "voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger",
    input_sliding,
    ids=sliding_ids
)
def test_demand_sliding(voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger):
    # 调用你的 generate_demand_sliding 函数
    generate_demand_sliding(voltage, angle, current, wire_type, demand_method, interval, update_rate, demand_trigger)


# ------------------- 脚本执行 -------------------
if __name__ == "__main__":
    try:
        pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html", '-x'])
    finally:
        if ModbusClient:
            ModbusClient.close()

