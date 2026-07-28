import time
import pyperclip
import pytest
from datetime import datetime
from comm.QT_comm.QT_utils.ModbusClient import ModbusProtocol, ModbusClient
from comm.source_control import *
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from comm.QT_comm.QT_utils.common_utils import CommonUtils
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
import struct
from modbus_config import modbus_config
import serial

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestTransaction:

    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        self.modbus_client = ModbusRtuOrTcp()
        # 初始化工具类
        self.utils = CommonUtils(self.helper, self.app_path, self.device_image_path)
        self.helper.kill_acuview_apps()
        self.helper.hotkey('win', 'd')
        self.helper.launch_app(self.app_path)
        yield
        self.helper.kill_acuview_apps()

    @pytest.mark.parametrize("config_value,expected_value,voltage,current", [
        (480, 480, 100, 500),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case1(self, config_value, expected_value, voltage, current):
        """配置交易timezone 为480,加直流源（100V，500A）,上位机下发”交易开始“命令"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Enable cable loss compensation', 0)
            client.validate_register_value('Cable  resistance', 0)
        self.utils.construct_transaction_logs(20, clear=False, connect=False)
        # 连接设备
        self.helper.connect_device(self.device_image_path)

        # 配置time_zone_shift（1110，363）
        self.utils.configure_time_zone_shift(value=config_value)
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle()
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 你的时区检查操作（保持不变）
        for i in actual_result:
            time_zone = i['TM']
            timezone_str = time_zone.split('+')[1].split()[0]  # "0800"
            hours = int(timezone_str[:2])
            minutes = int(timezone_str[2:4])
            total_minutes = hours * 60 + minutes

            assert total_minutes == expected_value, f"时区配置失败: 期望{expected_value}, 实际{total_minutes}"

        # 使用元素0和3计算能量和时间（单通道）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (kW)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / theory_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_time,voltage,current", [
        (60, 100, 500),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case2(self, config_time, voltage, current):
        """交易执行中，加直流源（100V，500A），本交易中的进线、出线电能精度满足0.5"""
        # 连接设备
        self.helper.connect_device(self.device_image_path)

        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和3计算能量和时间（进线能量）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_time,voltage,current", [
        (360, 100, -10),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case3(self, config_time, voltage, current):
        """交易执行中，加直流源（100V，-10A），本交易中的进线、出线电能精度满足0.5"""

        current = abs(current)  # 如果 current 已经是负数
        # 连接设备
        self.helper.connect_device(self.device_image_path)

        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和3计算能量和时间（出线能量）
        start_item = actual_result[1]  # 开始充电记录
        end_item = actual_result[3]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_cable, config_time,voltage,current", [
        (3, 60, 200, 10),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case4(self, config_cable, config_time, voltage, current):
        """交易执行中，加直流源（200V，10A），开启电缆损失补偿：3Ω，本交易中的进线、出线电能精度满足0.5"""

        self.modbus_client.write_registers(address=0X1022, values=[1], device_id=1)
        # 连接设备
        self.helper.connect_device(self.device_image_path)
        self.utils.configure_Cable(value='3')
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和3计算能量和时间（进线能量）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = (voltage - (current * config_cable)) * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_cable, config_time,voltage,current", [
        (3, 60, 200, -10),
    ])
    def test1_Function_AcuDC320_Sprint2_003_02_case5(self, config_cable, config_time, voltage, current):
        """交易执行中，加直流源（200V，-10A），开启电缆损失补偿：3Ω，本交易中的进线、出线电能精度满足0.5"""

        current = abs(current)  # 如果 current 已经是负数
        self.modbus_client.write_registers(address=0X1022, values=[1], device_id=1)
        # 连接设备
        self.helper.connect_device(self.device_image_path)
        time.sleep(5)
        self.utils.configure_Cable(value='3')
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和3计算能量和时间（进线能量）
        start_item = actual_result[1]  # 开始充电记录
        end_item = actual_result[3]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = (voltage + (current * config_cable)) * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    def test_Function_AcuDC320_Sprint2_003_02_case6(self):
        """交易执行中，再次触发”交易开始“命令，报错误信息或提示用户正在交易中，未开启新交易"""
        try:
            self.helper.connect_device(self.device_image_path)

            # 第一次点击开始充电（应该成功，不应该出现提示框）
            first_popup = self.utils.start_charging()
            assert not first_popup, "第一次点击开始充电时不应该出现'已经开始'提示框"

            # 第二次点击开始充电（应该出现提示框）
            second_popup = self.utils.start_charging()
            assert second_popup, "第二次点击开始充电时应该出现'已经开始'提示框"

        except Exception as e:
            print(f"测试执行过程中发生错误: {e}")
            raise
        finally:
            # 无论测试成功还是失败，都确保结束充电状态
            self.utils.end_charging()

    def test_Function_AcuDC320_Sprint2_003_02_case6_1(self):
        """未开始交易，上位机触发”结束/终止交易“命令“无效"""

        self.helper.connect_device(self.device_image_path)

        # 第一次点击结束充电（失败，出现提示框）
        first_popup = self.utils.end_charging()
        assert first_popup, "第一次点击结束充电时应该出现失败提示"

        # 第二次点击终止充电（失败，出现提示框）
        second_popup = self.utils.abort_charging()
        assert second_popup, "第二次点击结束充电时应该出现失败提示"

    def test_Function_AcuDC320_Sprint2_003_02_case6_2(self):
        """开始交易后，读取生成的交易日志，读取成功"""
        registers = self.modbus_client.read_measurement(0x5502, 2, device_id=1)
        # 组合成32位数：第一个寄存器是高16位，第二个是低16位,读取交易日志数量
        OCMF_count_old = (registers[0] << 16) | registers[1]
        self.helper.connect_device(self.device_image_path)
        # 开始充电-结束充电
        self.utils.perform_charging_cycle()
        time.sleep(2)
        registers = self.modbus_client.read_measurement(0x5502, 2, device_id=1)
        OCMF_count_new = (registers[0] << 16) | registers[1]
        assert OCMF_count_new == OCMF_count_old + 1, f"交易日志数量增加不正确。旧值: {OCMF_count_old}, 新值: {OCMF_count_new}"
        log_data = self.utils.read_and_parse_transaction_log()
        assert log_data is not None, "读取的交易日志数据为None"
        assert log_data != "", "读取的交易日志数据为空字符串"

    def test_Function_AcuDC320_Sprint2_003_02_case7(self):
        """交易执行中，上位机下发”交易结束“命令，交易结束并记录交易日志"""
        registers = self.modbus_client.read_measurement(0x5502, 2, device_id=1)
        # 组合成32位数：第一个寄存器是高16位，第二个是低16位,读取交易日志数量
        OCMF_count_old = (registers[0] << 16) | registers[1]
        self.helper.connect_device(self.device_image_path)
        # 开始充电
        self.utils.start_charging()
        # 查看充电后，变为charging状态
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_status_charging = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Charging')
        # 断言充电状态图片找到
        assert transaction_status_charging, "开始充电后应该显示Charging状态，但未找到Charging状态图片"
        # 结束充电
        self.utils.end_charging()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        # 开始充电后，已经点击过一次Transaction_In_Progress，标图为选中状态，充电结束后，需要点击其他按钮重置
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_status_idle = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Idle')
        # 断言空闲状态图片找到
        assert transaction_status_idle, "结束充电后应该显示Idle状态，但未找到Idle状态图片"
        registers = self.modbus_client.read_measurement(0x5502, 2, device_id=1)
        OCMF_count_new = (registers[0] << 16) | registers[1]
        assert OCMF_count_new == OCMF_count_old + 1, f"交易日志数量增加不正确。旧值: {OCMF_count_old}, 新值: {OCMF_count_new}"
        log_data = self.utils.read_and_parse_transaction_log()
        assert log_data is not None, "读取的交易日志数据为None"
        assert log_data != "", "读取的交易日志数据为空字符串"

    @pytest.mark.parametrize("config_time,voltage,current", [
        (60, 100, -500),
    ])
    def test1_Function_AcuDC320_Sprint2_003_02_case8(self, config_time, voltage, current):
        """交易日志精度验证，加直流源（100V，-500A），交易日志能量精度满足能量精度满足0.5"""
        # 连接设备
        current = abs(current)  # 如果 current 已经是负数
        self.helper.connect_device(self.device_image_path)
        self.modbus_client.write_registers(address=0X1022, values=[0], device_id=1)
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（进线能量）
        import_start_item = actual_result[0]  # 进线开始充电记录
        import_end_item = actual_result[2]  # 进线结束充电记录

        # 使用元素1和3计算能量和时间（出线能量）
        start_item = actual_result[1]  # 出线开始充电记录
        end_item = actual_result[3]  # 出线结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        import_start_energy = import_start_item['RV']  # 出线开始充电记录
        import_end_energy = import_end_item['RV']  # 出线结束充电记
        import_measured_energy = import_start_energy - import_end_energy
        assert import_measured_energy == 0
        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_time,voltage,current", [
        (60, 100, -500),
    ])
    def test1_Function_AcuDC320_Sprint2_003_02_case9(self, config_time, voltage, current):
        """交易日志精度验证，加直流源（100V，-500A），交易日志能量精度满足能量精度满足0.5"""
        self.modbus_client.write_registers(address=0X1022, values=[0], device_id=1)
        current = abs(current)  # 如果 current 已经是负数
        # 连接设备
        self.helper.connect_device(self.device_image_path)

        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（进线能量）
        import_start_item = actual_result[0]  # 进线开始充电记录
        import_end_item = actual_result[2]  # 进线结束充电记录

        # 使用元素1和3计算能量和时间（出线能量）
        start_item = actual_result[1]  # 出线开始充电记录
        end_item = actual_result[3]  # 出线结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        import_start_energy = import_start_item['RV']  # 出线开始充电记录
        import_end_energy = import_end_item['RV']  # 出线结束充电记
        import_measured_energy = import_start_energy - import_end_energy
        assert import_measured_energy == 0
        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_cable, config_time,voltage,current", [
        (3, 60, 100, 50),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case10(self, config_cable, config_time, voltage, current):
        """交易日志精度验证，加直流源（100V，50A），开启电缆损失补偿：3Ω，交易日志能量精度满足能量精度满足0.5"""

        resp = self.modbus_client.write_registers(address=0X1022, values=[1], device_id=1)
        # 连接设备
        self.helper.connect_device(self.device_image_path)
        self.utils.configure_Cable(value='3')
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（进线能量）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = (voltage - (current * config_cable)) * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_cable, config_time,voltage,current", [
        (3, 60, 200, -100),
    ])
    def test1_Function_AcuDC320_Sprint2_003_02_case11(self, config_cable, config_time, voltage, current):
        """交易日志精度验证，加直流源（100V，-500A），开启电缆损失补偿：3Ω，交易日志能量精度满足能量精度满足0.5"""

        current = abs(current)  # 如果 current 已经是负数
        resp = self.modbus_client.write_registers(address=0X1022, values=[1], device_id=1)
        # 连接设备
        self.helper.connect_device(self.device_image_path)
        self.utils.configure_Cable(value='3')
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（出现线能量）
        start_item = actual_result[1]  # 开始充电记录
        end_item = actual_result[3]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = (voltage + (current * config_cable)) * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_cable, config_time,voltage,current", [
        (3, 60, 200, -100),
    ])
    def test1_Function_AcuDC320_Sprint2_003_02_case11_1(self, config_cable, config_time, voltage, current):
        """交易记录数据准确性验证。交易开始和结束点的能量寄存器与交易日志中记录是否一致"""

        current = abs(current)  # 如果 current 已经是负数
        self.modbus_client.write_registers(address=0X1022, values=[1], device_id=1)
        # 连接设备
        self.helper.connect_device(self.device_image_path)
        self.utils.configure_Cable(value='3')
        # 读取Modbus寄存器进线能量值
        import_registers = self.modbus_client.read_measurement(0x4000, 4, device_id=1)
        export_registers = self.modbus_client.read_measurement(0x4004, 4, device_id=1)
        # 转换为浮点数并格式化为4位小数
        import_data_bytes = b''.join(reg.to_bytes(2, 'big') for reg in import_registers)
        import_float_value = struct.unpack('>d', import_data_bytes)[0]
        import_value_old = f"{import_float_value:.4f}"

        export_data_bytes = b''.join(reg.to_bytes(2, 'big') for reg in export_registers)
        export_float_value = struct.unpack('>d', export_data_bytes)[0]
        export_value_old = f"{export_float_value:.4f}"

        print(f"读取进线能量的值: {import_value_old} kWh")
        print(f"读取出现能量的值: {export_value_old} kWh")
        sour_output(voltage, current)
        # 执行充电
        actual_start_time, actual_end_time = self.utils.perform_charging_cycle(wait=config_time)
        sour_stop()

        # 读取Modbus寄存器进线能量值
        import_registers = self.modbus_client.read_measurement(0x4000, 4, device_id=1)
        export_registers = self.modbus_client.read_measurement(0x4004, 4, device_id=1)
        # 转换为浮点数并格式化为4位小数
        import_data_bytes = b''.join(reg.to_bytes(2, 'big') for reg in import_registers)
        import_float_value = struct.unpack('>d', import_data_bytes)[0]
        import_value_new = f"{import_float_value:.4f}"

        export_data_bytes = b''.join(reg.to_bytes(2, 'big') for reg in export_registers)
        export_float_value = struct.unpack('>d', export_data_bytes)[0]
        export_value_new = f"{export_float_value:.4f}"
        print(f"读取充电后进线能量的值: {import_value_new} kWh")
        print(f"读取充电后出现能量的值: {export_value_new} kWh")
        # 计算实际充电时长
        actual_duration = actual_end_time - actual_start_time
        actual_start_datetime = datetime.fromtimestamp(actual_start_time)
        actual_end_datetime = datetime.fromtimestamp(actual_end_time)

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（出现线能量）
        import_start_item = actual_result[0]  # 开始充电记录
        import_end_item = actual_result[2]  # 结束充电记录

        # 使用元素1和3计算能量和时间（进线线能量）
        start_item = actual_result[1]  # 开始充电记录
        end_item = actual_result[3]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"实际开始时间: {actual_start_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志开始时间: {start_time_clean}")
        print(f"实际结束时间: {actual_end_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"实际充电时长: {actual_duration:.2f}秒")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        time_tolerance = 5  # 5秒容差
        actual_vs_log_start_diff = abs((actual_start_datetime - log_start_time).total_seconds())
        actual_vs_log_end_diff = abs((actual_end_datetime - log_end_time).total_seconds())
        duration_diff = abs(actual_duration - log_duration)

        assert actual_vs_log_start_diff <= time_tolerance, f"开始时间差异过大: {actual_vs_log_start_diff}秒"
        assert actual_vs_log_end_diff <= time_tolerance, f"结束时间差异过大: {actual_vs_log_end_diff}秒"
        assert duration_diff <= time_tolerance, f"充电时长差异过大: {duration_diff}秒"

        actual_duration = actual_end_time - actual_start_time
        # 存在误差时间，开始输出时间必定大于实际充电时间
        real_time = actual_duration - log_duration
        # 计算理论能量（根据你的输入电压电流）
        power = (voltage + (current * config_cable)) * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量
        # 断言实测能量等于出线能量差值（考虑浮点数精度）
        export_energy_diff = float(export_value_new) - float(export_value_old)
        assert abs(measured_energy - export_energy_diff) < real_time * power, (
            f"实测能量与出线能量不匹配。实测: {measured_energy:.6f}, "
            f"出线能量差: {export_energy_diff:.6f},"
            f"误差值为{real_time * power}"
        )
        # 断言进线能量差值为0（进线能量应该没有变化）
        import_energy_diff = float(import_value_new) - float(import_value_old)
        assert abs(import_energy_diff) <= 0.001, (
            f"进线能量误差不变，但实际变化为: {import_energy_diff:.6f}"
        )

        print("✓ 能量验证通过：实测能量与出线能量一致，进线能量无变化")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

    @pytest.mark.parametrize("config_time,voltage,current", [
        (60, 100, 500),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case12(self, config_time, voltage, current):
        """交易过程中，异常断电，能结束并记录交易日志，能量精度满足能量精度满足0.5"""

        # 连接设备
        self.helper.connect_device(self.device_image_path)
        sour_output(voltage, current)
        # 执行充电
        com = modbus_config.modbus_config['rtu']['port']
        baudrate = modbus_config.modbus_config['rtu']['baudrate']
        reboot_cmd = bytes([0x01, 0x6A, 0xF0, 0x00, 0x00, 0x02, 0x04, 0xA1, 0x58, 0x58, 0xA1, 0x8D, 0xF7])
        self.utils.start_charging()
        time.sleep(config_time)
        self.modbus_client.close()

        with serial.Serial(com, baudrate, timeout=2) as ser:
            ser.write(reboot_cmd)
            print("重启命令已发送")
        sour_stop()

        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和2计算能量和时间（进线能量）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 使用元素1和3计算能量和时间（出线能量）
        export_start_item = actual_result[0]  # 开始充电记录
        export_end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息

        print(f"日志开始时间: {start_time_clean}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        print(f"日志充电时长: {log_duration:.2f}秒, 配置充电时长: {config_time}秒")

        # 断言日志充电时长与配置时长的差异不超过5秒
        config_time_diff = abs(log_duration - config_time)
        assert config_time_diff <= 5, f"日志时长与配置时长差异过大: {config_time_diff:.2f}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

        assert export_end_item['TX'] == 'P', f"结束充电标识符错误，期望: 'P', 实际: {export_end_item['TX']}"
        assert end_item['TX'] == 'P', f"结束充电标识符错误，期望: 'P', 实际: {end_item['TX']}"

    @pytest.mark.parametrize(" config_time,voltage,current", [
        (3600, 100, 500),
    ])
    def test_Function_AcuDC320_Sprint2_003_02_case13(self, config_time, voltage, current):
        """交易过程中，异常超时，能结束并记录交易日志，能量精度满足能量精度满足0.5"""

        # 连接设备
        self.helper.connect_device(self.device_image_path)
        self.utils.configure_Transaction_Timeout(config_time)
        sour_output(voltage, current)
        # 执行充电
        self.utils.long_time_charging(config_time)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        sour_stop()
        # 读取并解析日志
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']

        # 使用元素0和3计算能量和时间（进线能量）
        start_item = actual_result[0]  # 开始充电记录
        end_item = actual_result[2]  # 结束充电记录

        # 解析日志中的时间
        start_time_str = start_item['TM'].split('+')[0]  # "2024-01-01T03:38:16,955"
        end_time_str = end_item['TM'].split('+')[0]  # "2024-01-01T03:38:22,063"

        # 解析日志时间（去除毫秒）
        start_time_clean = start_time_str.split(',')[0]  # "2024-01-01T03:38:16"
        end_time_clean = end_time_str.split(',')[0]  # "2024-01-01T03:38:22"

        log_start_time = datetime.strptime(start_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_end_time = datetime.strptime(end_time_clean, "%Y-%m-%dT%H:%M:%S")
        log_duration = (log_end_time - log_start_time).total_seconds()

        # 打印时间信息
        print(f"日志开始时间: {start_time_clean}")
        print(f"日志结束时间: {end_time_clean}")
        print(f"日志充电时长: {log_duration:.2f}秒")

        # 时间断言比较（允许一定的时间误差）
        print(f"日志充电时长: {log_duration:.2f}秒, 配置充电时长: {config_time}秒")

        # 断言日志充电时长与配置时长的差异不超过5秒
        config_time_diff = abs(log_duration - config_time)
        assert config_time_diff <= 5, f"日志时长与配置时长差异过大: {config_time_diff:.2f}秒"

        # 计算理论能量（根据你的输入电压电流）
        power = voltage * current * 0.001  # 功率 (W)
        theory_energy = power * (log_duration / 3600)  # 能量 (Wh)，使用实际时长

        print(f"输入功率: {power}kW")
        print(f"理论能量: {theory_energy:.6f} kWh")

        # 能量验证（如果RV字段是能量值）
        start_energy = start_item['RV']  # 开始时的能量读数
        end_energy = end_item['RV']  # 结束时的能量读数
        measured_energy = end_energy - start_energy  # 实际消耗的能量

        print(f"开始能量: {start_energy} {start_item['RU']}")
        print(f"结束能量: {end_energy} {end_item['RU']}")
        print(f"实测能量: {measured_energy} {end_item['RU']}")

        # 能量断言比较（处理能量为0的情况）
        energy_tolerance = 0.005  # 0.5%误差
        if measured_energy == 0:
            # 如果实测能量为0，检查理论能量是否也很小（接近0）
            assert theory_energy <= 0.001, f"能量应该为0，但理论能量为{theory_energy:.6f} Wh"
            print("能量为0，验证通过")
        else:
            # 如果实测能量不为0，使用相对误差验证
            energy_error = abs(theory_energy - measured_energy) / measured_energy
            assert energy_error <= energy_tolerance, f"能量误差过大: {energy_error * 100:.2f}%"
            print(f"能量验证通过，误差: {energy_error * 100:.2f}%")

        assert end_item['TX'] == 'P', f"结束充电标识符错误，期望: 'P', 实际: {end_item['TX']}"

    def test_Function_AcuDC320_Sprint2_003_02_case13_1(self):
        """结束交易后，Transaction In Progress中电能和充电时长需要重置为“-”"""

        self.helper.connect_device(self.device_image_path)
        # 开始充电
        self.utils.start_charging()
        # 查看充电后，变为charging状态
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_satrt_time = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_satrt_time')
        # 断言充电状态图片找到
        assert not transaction_satrt_time, "开始充电后充电开始时间不为默认值"
        # 结束充电
        self.utils.end_charging()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        # 开始充电后，已经点击过一次Transaction_In_Progress，标图为选中状态，充电结束后，需要点击其他按钮重置
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_end_time = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_satrt_time')
        # 结束充电后充电开始时间为默认值
        assert transaction_end_time, "结束充电后应该显示Idle状态，但未找到Idle状态图片"

    def test_Function_AcuDC320_Sprint2_003_02_case13_2(self):
        """开始交易后，Transaction In Progress中电能和充电时长需要重置为发送交易时间点”"""

        self.helper.connect_device(self.device_image_path)
        # 开始充电
        self.utils.start_charging()
        time.sleep(10)
        # 查看充电后，变为charging状态
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_satrt_time = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_satrt_time')
        transaction_duration_time = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_duration_time')

        assert not transaction_satrt_time, "开始充电后充电开始时间不为默认值"
        assert not transaction_duration_time, "开始充电后充电总时长不为默认值"
        # 结束充电
        self.utils.end_charging()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        # 开始充电后，已经点击过一次Transaction_In_Progress，标图为选中状态，充电结束后，需要点击其他按钮重置
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_end_time = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_satrt_time')
        transaction_end_duration = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Transaction_duration_time')
        # 断言充电状态图片找到
        # 结束充电后充电开始时间为默认值
        assert transaction_end_time, "结束充电后应该显示Idle状态，但未找到Idle状态图片"
        assert transaction_end_duration, "结束充电后应该显示Idle状态，但未找到Idle状态图片"

    def test_Function_AcuDC320_Sprint2_003_02_case13_3(self):
        """电表下电后上电，Transaction In Progress中时间同步更新为 No sync”"""

        # 连接设备
        self.helper.connect_device(self.device_image_path)
        com = modbus_config.modbus_config['rtu']['port']
        baudrate = modbus_config.modbus_config['rtu']['baudrate']
        reboot_cmd = bytes([0x01, 0x6A, 0xF0, 0x00, 0x00, 0x02, 0x04, 0xA1, 0x58, 0x58, 0xA1, 0x8D, 0xF7])

        self.utils.configure_time_Sync_status()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_time_sync = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Sync')
        self.modbus_client.close()
        assert transaction_time_sync, f"显示时间同步状态不正确，预期为'Sync'，实际为{transaction_time_sync}"
        with serial.Serial(com, baudrate, timeout=2) as ser:
            ser.write(reboot_cmd)
            print("重启命令已发送")
        time.sleep(20)
        self.utils.restart_application()
        self.helper.connect_device(self.device_image_path)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading', 0)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_In_Progress')
        transaction_time_not_sync = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Reading_page\Transaction_in_Progress\Sync')
        assert not transaction_time_not_sync, f"显示时间同步状态不正确，预期为'not Sync'，实际为{transaction_time_sync}"


if __name__ == "__main__":
    # 直接运行测试
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html"])
