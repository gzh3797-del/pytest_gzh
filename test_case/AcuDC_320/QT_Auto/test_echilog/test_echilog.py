import time
import pyperclip
import pytest
import sys
import os
from comm.QT_comm.QT_utils.ModbusClient import ModbusClient, ModbusProtocol
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from comm.QT_comm.QT_utils.common_utils import CommonUtils
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from modbus_config import modbus_config

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestEchilog:
    #
    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.app_root_path = os.path.dirname(self.app_path)
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        self.modbus_client = ModbusRtuOrTcp()
        # 初始化工具类
        self.utils = CommonUtils(self.helper, self.app_path, self.device_image_path)
        self.helper.kill_acuview_apps()
        self.helper.wait(2)
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        yield
        self.helper.kill_acuview_apps()

    def test_Function_AcuDC320_Sprint2_003_02_case1(self):
        """系统状态只读：通过Modbus寄存器可读，不可写入"""

        read_result = self.modbus_client.read_measurement(address=0XF0B6, count=1, slave=1)

        write_result = self.modbus_client.write_registers(address=0XF0B6, values=[1], slave=1)
        # 断言读取有返回值
        assert read_result and len(read_result) > 0, "读取操作应返回有效数据"

        # 断言：write不允许写入（返回错误/异常即可）
        assert write_result is None or hasattr(write_result, 'exception_code'), "写入操作应返回错误或异常"

    def test_Function_AcuDC320_Sprint2_002_04_case1(self):
        """1、Cable Loss Enable配置：Disable-Enable，生成一条Echilog，记录旧新值、时间戳、ID"""

        self.modbus_client.write_registers(address=0X1022, values=[0], slave=1)

        self.helper.connect_device(self.device_image_path)
        old_echilog = self.modbus_client.read_measurement(address=0X5703, count=1, slave=1)
        self.utils.configure_Cable_status(value='1')
        new_echilog = self.modbus_client.read_measurement(address=0X5703, count=1, slave=1)
        assert new_echilog[0] == old_echilog[0] + 1, f"f'读取echilog日志不正确，{old_echilog},{new_echilog}'"
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        # 新增断言
        assert type_value == "Cable Loss Enable Status Change", f"类型不正确，期望: Cable Loss Enable Status Change, 实际: {type_value}"
        assert old_value == "Disable", f"旧值不正确，期望: Disable, 实际: {old_value}"
        assert new_value == "Enable", f"新值不正确，期望: Enable, 实际: {new_value}"

    def test_Function_AcuDC320_Sprint2_002_04_case2(self):
        """Cable Loss Enable配置：Enable-Disable，生成一条Echilog，记录旧新值、时间戳、ID"""
        self.modbus_client.write_registers(address=0X1022, values=[1], slave=1)
        self.helper.connect_device(self.device_image_path)
        old_echilog = self.modbus_client.read_measurement(address=0X5703, count=1, slave=1)
        self.utils.configure_Cable_status(value='0')
        new_echilog = self.modbus_client.read_measurement(address=0X5703, count=1, slave=1)
        assert new_echilog[0] == old_echilog[0] + 1, f"f'读取echilog日志不正确，{old_echilog},{new_echilog}'"
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        # 新增断言
        assert type_value == "Cable Loss Enable Status Change", f"类型不正确，期望: Cable Loss Enable Status Change, 实际: {type_value}"
        assert old_value == "Enable", f"旧值不正确，期望: Disable, 实际: {old_value}"
        assert new_value == "Disable", f"新值不正确，期望: Enable, 实际: {new_value}"

    def test_Function_AcuDC320_Sprint2_002_04_case3(self):
        """3、Cable Loss Resistance配置：0.0001Ω，生成一条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Enable cable loss compensation', 1)
            client.validate_register_value('Cable  resistance', 0)
        self.utils.connect_device()
        self.utils.configure_cable_loss(resistance=0.0001)
        result = self.utils.read_echilog_multi_line(1)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        pytest.assume(type_value1 == "Cable Loss Resistanche Change",
                      f"类型不正确，期望: Cable Loss Resistanche Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "0.0000", f"旧值不正确，期望: 0.0000, 实际: {old_value1}")
        pytest.assume(new_value1 == "0.0001", f"新值不正确，期望: 0.0001, 实际: {new_value1}")

    def test_Function_AcuDC320_Sprint2_002_04_case4(self):
        """Cable Loss Resistance配置：3Ω，生成一条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Enable cable loss compensation', 1)
            client.validate_register_value('Cable  resistance', 0)
        self.utils.connect_device()
        self.utils.configure_cable_loss(resistance=3)
        result = self.utils.read_echilog_multi_line(1)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        pytest.assume(type_value1 == "Cable Loss Resistanche Change",
                      f"类型不正确，期望: Cable Loss Resistanche Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "0.0000", f"旧值不正确，期望: 0.0000, 实际: {old_value1}")
        pytest.assume(new_value1 == "3.0000", f"新值不正确，期望: 3.0000, 实际: {new_value1}")

    def test_Function_AcuDC320_Sprint2_002_04_case5(self):
        """Cable Loss Resistance配置：6.5535Ω，生成一条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Enable cable loss compensation', 1)
            client.validate_register_value('Cable  resistance', 0)
        self.utils.connect_device()
        self.utils.configure_cable_loss(resistance=6.5535)
        result = self.utils.read_echilog_multi_line(1)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        pytest.assume(type_value1 == "Cable Loss Resistanche Change",
                      f"类型不正确，期望: Cable Loss Resistanche Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "0.0000", f"旧值不正确，期望: 0.0000, 实际: {old_value1}")
        pytest.assume(new_value1 == "6.5535", f"新值不正确，期望: 6.5535, 实际: {new_value1}")

    def test_Function_AcuDC320_Sprint2_002_04_case6(self):
        """同时修改Cable Loss Enable和Cable Loss Resistance配置：Enable、3Ω，生成两条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Enable cable loss compensation', 0)
            client.validate_register_value('Cable  resistance', 0)
        self.utils.connect_device()
        self.utils.configure_cable_loss(status='on', resistance=3)
        result = self.utils.read_echilog_multi_line(2)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        time_value2, type_value2, old_value2, new_value2 = result[1]
        pytest.assume(type_value1 == "Cable Loss Enable Status Change",
                      f"类型不正确，期望: Cable Loss Enable Status Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "Disable", f"旧值不正确，期望: Disable, 实际: {old_value1}")
        pytest.assume(new_value1 == "Enable", f"新值不正确，期望: Enable, 实际: {new_value1}")
        pytest.assume(type_value2 == "Cable Loss Resistanche Change",
                      f"类型不正确，期望: Cable Loss Resistanche Change, 实际: {type_value2}")
        pytest.assume(old_value2 == "0.0000", f"旧值不正确，期望:0.0000, 实际: {old_value2}")
        pytest.assume(new_value2 == "3.0000", f"新值不正确，期望: 3.0000, 实际: {new_value2}")

    def test_Function_AcuDC320_Sprint2_002_05_case1(self):
        """
        Pulse LED Energy配置改变：Energy pulse parameter： None—Import Energy ，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Energy pulse parameter', 0)
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='Import_Energy')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value}")
        pytest.assume(old_value == "None", f"旧值不正确，期望: None, 实际: {old_value}")
        pytest.assume(new_value == "Import Energy", f"新值不正确，期望: Import Energy, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case2(self):
        """
        Pulse LED Energy配置改变：Energy pulse parameter：Import Energy —Export Energy  ，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='Export_Energy')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value}")
        pytest.assume(old_value == "Import Energy", f"旧值不正确，期望: Import Energy, 实际: {old_value}")
        pytest.assume(new_value == "Export Energy", f"新值不正确，期望: Export Energy, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case3(self):
        """
        Pulse LED Energy配置改变：Energy pulse parameter：Export Energy  — NET Energy，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='NET_Energy')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value}")
        pytest.assume(old_value == "Export Energy", f"旧值不正确，期望: Export Energy, 实际: {old_value}")
        pytest.assume(new_value == "NET Energy", f"新值不正确，期望: NET Energy, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case4(self):
        """
        Pulse LED Energy配置改变：Energy pulse parameter： NET Energy  —TOTAL Energy，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='TOTAL_Energy')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value}")
        pytest.assume(old_value == "NET Energy", f"旧值不正确，期望: NET Energy, 实际: {old_value}")
        pytest.assume(new_value == "TOTAL Energy", f"新值不正确，期望: TOTAL Energy, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case5(self):
        """
        Pulse LED Energy配置改变：Energy pulse parameter：TOTAL Energ — None，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='None')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value}")
        pytest.assume(old_value == "TOTAL Energy", f"旧值不正确，期望: TOTAL Energy, 实际: {old_value}")
        pytest.assume(new_value == "None", f"新值不正确，期望: None, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case6(self):
        """
        Pulse LED Energy配置改变：Energy pulse constant：0.1，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Energy pulse constant', 100000)
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(constant=0.1)
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Pulse Change",
                      f"类型不正确，期望: Energy Pulse Change, 实际: {type_value}")
        pytest.assume(old_value == "100.000", f"旧值不正确，期望: 100.000, 实际: {old_value}")
        pytest.assume(new_value == "0.100", f"新值不正确，期望: 0.100, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case7(self):
        """
        Pulse LED Energy配置改变：Energy pulse constant：1000，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(constant=1000)
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Pulse Change",
                      f"类型不正确，期望: Energy Pulse Change, 实际: {type_value}")
        pytest.assume(old_value == "0.100", f"旧值不正确，期望: 0.100, 实际: {old_value}")
        pytest.assume(new_value == "1000.000", f"新值不正确，期望: 1000.000, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case8(self):
        """
        Pulse LED Energy配置改变：Energy pulse constant：100000，
        生成一条Echilog，记录旧新值、时间戳、ID
        """
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(constant=100000)
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Energy Pulse Change",
                      f"类型不正确，期望: Energy Pulse Change, 实际: {type_value}")
        pytest.assume(old_value == "1000.000", f"旧值不正确，期望:1000.000, 实际: {old_value}")
        pytest.assume(new_value == "100000.000", f"新值不正确，期望: 100000.000, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_05_case8_1(self):
        """
        同时修改Pulse LED Energy配置：Energy pulse constant：100000、Energy pulse parameter： None—Import Energy ，
        生成两条Echilog，记录旧新值、时间戳、ID
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Energy pulse constant', 100000)
            client.validate_register_value('Energy pulse parameter', 0)
        self.utils.connect_device()
        self.utils.configure_Pulse_LED_Energy(parameter='Import_Energy', constant=100000)
        result = self.utils.read_echilog_multi_line(2)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        time_value2, type_value2, old_value2, new_value2 = result[1]
        pytest.assume(type_value1 == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "None", f"旧值不正确，期望: None, 实际: {old_value1}")
        pytest.assume(new_value1 == "Import Energy", f"新值不正确，期望: Import Energy, 实际: {new_value1}")
        pytest.assume(type_value2 == "Energy Pulse Change",
                      f"类型不正确，期望: Energy Pulse Change, 实际: {type_value2}")
        pytest.assume(old_value2 == "100.000", f"旧值不正确，期望:100.000, 实际: {old_value2}")
        pytest.assume(new_value2 == "100000.000", f"新值不正确，期望: 100000.000, 实际: {new_value2}")

    def test_Function_AcuDC320_Sprint2_002_05_case8_2(self):
        """
        同时修改Pulse LED Energy配置与Cable Loss配置， 生成两条Echilog，记录旧新值、时间戳、ID
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Energy pulse parameter', 0)
            client.validate_register_value('Enable cable loss compensation', 0)
        self.utils.connect_device()
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
        # 点击General
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable_disable', offset_x=-66)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\parameter', offset_x=211)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\parameter', offset_x=127,
                                offset_y=39)
        self.utils._update_configuration()
        # 断言检查结果
        result = self.utils.read_echilog_multi_line(2)
        time_value1, type_value1, old_value1, new_value1 = result[0]
        time_value2, type_value2, old_value2, new_value2 = result[1]
        pytest.assume(type_value1 == "Energy Parameter Change",
                      f"类型不正确，期望: Energy Parameter Change, 实际: {type_value1}")
        pytest.assume(old_value1 == "None", f"旧值不正确，期望: None, 实际: {old_value1}")
        pytest.assume(new_value1 == "Import Energy", f"新值不正确，期望: Import Energy, 实际: {new_value1}")
        pytest.assume(type_value2 == "Cable Loss Enable Status Change",
                      f"类型不正确，期望: Cable Loss Enable Status Change, 实际: {type_value2}")
        pytest.assume(old_value2 == "Disable", f"旧值不正确，期望:Disable, 实际: {old_value2}")
        pytest.assume(new_value2 == "Enable", f"新值不正确，期望: Enable, 实际: {new_value2}")

    def test_Function_AcuDC320_Sprint2_002_06_case1(self):
        """Charge Point Identification Type配置变化：0: EVSEID  —1: CBIDC，
        生成一条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('CT: Charge Point Identification Type', 0)
        self.utils.connect_device()
        self.utils.configure_CT(value='CBIDC')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Charge Point IT Change",
                      f"类型不正确，期望: Charge Point IT Change, 实际: {type_value}")
        pytest.assume(old_value == "EVSEID", f"旧值不正确，期望: EVSEID, 实际: {old_value}")
        pytest.assume(new_value == "CBIDC", f"新值不正确，期望: CBIDC, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_06_case2(self):
        """Charge Point Identification Type配置变化： 1: CBIDC—0: EVSEID ，
        生成一条Echilog，记录旧新值、时间戳、ID"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('CT: Charge Point Identification Type', 1)
        self.utils.connect_device()
        self.utils.configure_CT(value='EVSEID')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Charge Point IT Change",
                      f"类型不正确，期望: Charge Point IT Change, 实际: {type_value}")
        pytest.assume(old_value == "CBIDC", f"旧值不正确，期望: CBIDC, 实际: {old_value}")
        pytest.assume(new_value == "EVSEID", f"新值不正确，期望: EVSEID, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_06_case3(self):
        """Charge Point Identification配置变化：20个寄存器（默认全0），改变任意寄存器值，
        生成一条Echilog，记录旧新值、时间戳、ID"""
        self.utils.connect_device()
        self.utils.configure_CI(value='NONE')
        self.utils.restart_application()
        self.utils.connect_device()
        self.utils.configure_CI(value='@@@@!!!!!%%%&%%%$1122sdasdvvfad')
        time_value, type_value, old_value, new_value = self.utils.read_echilog_info()
        pytest.assume(type_value == "Charge Point ID Change",
                      f"类型不正确，期望: Charge Point ID Change, 实际: {type_value}")
        pytest.assume(old_value == "NONE", f"旧值不正确，期望: NONE, 实际: {old_value}")
        pytest.assume(new_value == "@@@@!!!!!%%%&%%%$1122sdasdvvfad",
                      f"新值不正确，期望: @@@@!!!!!%%%&%%%$1122sdasdvvfad, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_07_case1(self):
        """时间同步相差小于300秒（小于300秒）， 未生成一条Echilog"""
        self.utils.connect_device()
        old_time, new_time = self.utils.set_time(60, 'Reading')
        result = self.utils.read_echilog_multi_line(1)
        time_value, type_value, old_value, new_value = result[0]
        pytest.assume(type_value != "RTC Configuration Mismatch",
                      f"类型不正确，期望: RTC Configuration Mismatch, 实际: {type_value}")
        pytest.assume(old_value != old_time,
                      f"旧值不正确，期望: {old_time}, 实际: {old_value}")
        pytest.assume(new_value != new_time,
                      f"新值不正确，期望: {new_time}, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_07_case2(self):
        """时间同步相差大于300秒（大于300秒）， 生成一条Echilog，记录旧新值、时间戳、ID"""
        self.utils.connect_device()
        old_time, new_time = self.utils.set_time(305, 'Reading')
        result = self.utils.read_echilog_multi_line(1)
        time_value, type_value, old_value, new_value = result[0]
        pytest.assume(type_value == "RTC Configuration Mismatch",
                      f"类型不正确，期望: RTC Configuration Mismatch, 实际: {type_value}")
        pytest.assume(self.utils.is_time_within_2_seconds(old_value, old_time),
                      f"旧值不正确，期望: {old_time}, 实际: {old_value}")
        pytest.assume(self.utils.is_time_within_2_seconds(new_value, new_time),
                      f"新值不正确，期望: {new_time}, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_07_case3(self):
        """时间同步相差大于300秒（大于300秒）， 生成一条Echilog，记录旧新值、时间戳、ID"""
        self.utils.connect_device()
        old_time, new_time = self.utils.set_time(340, 'Setting')
        result = self.utils.read_echilog_multi_line(1)
        time_value, type_value, old_value, new_value = result[0]
        pytest.assume(type_value == "RTC Configuration Mismatch",
                      f"类型不正确，期望: RTC Configuration Mismatch, 实际: {type_value}")
        pytest.assume(self.utils.is_time_within_2_seconds(old_value, old_time),
                      f"旧值不正确，期望: {old_time}, 实际: {old_value}")
        pytest.assume(self.utils.is_time_within_2_seconds(new_value, new_time),
                      f"新值不正确，期望: {new_time}, 实际: {new_value}")

    def test_Function_AcuDC320_Sprint2_002_08_case10(self):
        """铅封打开，清除Echilog记录。清除成功"""
        self.utils.construct_EClig_logs(50, clear=False)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\clear_log')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        time.sleep(1)
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_Log\clear_successful'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        else:
            pytest.fail('交易日志清理失败')
        current_log_num = int(self.helper.quick_ocr_by_config('EC Used Records'))
        assert current_log_num == 0, 'Echilog提示清理失败，Used Records值不为0'

    @pytest.mark.parametrize("config_value,expected_value", [
        (49, 49),
        (50, 50),
        (51, 50)
    ])
    def test_Function_AcuDC320_Sprint2_002_08_case2(self, config_value, expected_value):
        """读取Echilog，上位机选择 Read latest 50 Records 读取正常"""
        self.utils.construct_EClig_logs(config_value)
        self.utils.read_logs()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.hotkey('ctrl', 'c')
        # 获取文件名
        file_name = pyperclip.paste()
        self.helper.hotkey('enter')
        time.sleep(5)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        file_path = rf'{self.app_root_path}\Export\{file_name}'
        expected_data = self.utils.read_echilog_info()
        self.helper.check_csv_file(file_path, expected_value, expected_data, 'EC')

    @pytest.mark.parametrize("config_value,expected_value", [
        (999, 999),
        (1000, 1000),
        (1001, 1000),
    ])
    def test_Function_AcuDC320_Sprint2_002_08_case3(self, config_value, expected_value):
        """读取Echilog，上位机选择Read latest 1000 Records读取正常"""
        self.utils.construct_EClig_logs(config_value)
        self.helper.click_pos((1005, 208))
        self.helper.click_pos((928, 245))
        self.utils.read_logs()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.hotkey('ctrl', 'c')
        # 获取文件名
        file_name = pyperclip.paste()
        self.helper.hotkey('enter')
        time.sleep(5)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        file_path = rf'{self.app_root_path}\Export\{file_name}'
        expected_data = self.utils.read_echilog_info()
        self.helper.check_csv_file(file_path, expected_value, expected_data, 'EC')

    @pytest.mark.parametrize("config_value", [
        '1',
        '10',
        '5999',
        '6001'
    ])
    def test_Function_AcuDC320_Sprint2_002_08_case4(self, config_value):
        """读取Echilog，上位机选择Read latest 500 Records(from selected Record)读取正常"""
        self.utils.construct_EClig_logs(6000, clear=False)
        self.helper.click_pos((937, 208))
        self.helper.click_pos((893, 263))
        self.helper.click_pos((1183, 211))
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(config_value)
        if config_value == '6001':
            self.helper.click_pos((1120, 272))
            Start_Id = int(self.helper.quick_ocr_by_config('EC Start Id'))
            if Start_Id == int(config_value):
                pytest.fail('Start_Id配置6001成功')
        else:
            self.utils.read_logs()
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.hotkey('ctrl', 'c')
            # 获取文件名
            file_name = pyperclip.paste()
            self.helper.hotkey('enter')
            time.sleep(5)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            file_path = rf'{self.app_root_path}\Export\{file_name}'
            used = int(self.helper.quick_ocr_by_config('EC Used Records'))
            read_num = used - int(config_value) + 1
            if read_num > 500:
                expected_data_len = 500
            else:
                expected_data_len = read_num
            expected_data = self.utils.read_echilog_info()
            self.helper.check_csv_file(file_path, expected_data_len, expected_data, 'EC')

    @pytest.mark.parametrize("config_value", [
        '1',
        '10',
        '5999',
        '6001'
    ])
    def test_Function_AcuDC320_Sprint2_002_08_case5(self, config_value):
        """读取Echilog，上位机选择Read latest 1000 Records(from selected Record)读取正常"""
        self.utils.construct_EClig_logs(6000, clear=False)
        self.helper.click_pos((937, 208))
        self.helper.click_pos((878, 278))
        self.helper.click_pos((1183, 211))
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(config_value)
        if config_value == '6001':
            self.helper.click_pos((1120, 272))
            Start_Id = int(self.helper.quick_ocr_by_config('EC Start Id'))
            if Start_Id == int(config_value):
                pytest.fail('Start_Id配置6001成功')
        else:
            self.utils.read_logs()
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.hotkey('ctrl', 'c')
            # 获取文件名
            file_name = pyperclip.paste()
            self.helper.hotkey('enter')
            time.sleep(5)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            file_path = rf'{self.app_root_path}\Export\{file_name}'
            used = int(self.helper.quick_ocr_by_config('EC Used Records'))
            read_num = used - int(config_value) + 1
            if read_num > 1000:
                expected_data_len = 1000
            else:
                expected_data_len = read_num
            expected_data = self.utils.read_echilog_info()
            self.helper.check_csv_file(file_path, expected_data_len, expected_data, 'EC')

    @pytest.mark.parametrize("config_value,expected_value", [
        (50, 50),
        (1000, 1000),
    ])
    def test_Function_AcuDC320_Sprint2_002_08_case6(self, config_value, expected_value):
        """上位机读取日志，循环执行读取50条、1000条日志20次，读取正常"""
        self.utils.construct_EClig_logs(config_value, clear=False)
        for i in range(20):
            if config_value == 50:
                self.utils.read_logs()
            else:
                self.helper.click_pos((1005, 208))
                self.helper.click_pos((928, 245))
                self.utils.read_logs()

    def test_Function_AcuDC320_Sprint2_002_08_case7(self):
        """上位机读取日志时，点击停止按钮，可停止日志读取"""
        self.utils.construct_EClig_logs(20, clear=False)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        time.sleep(2)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Stop')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Failed'):
            pytest.fail('读取日志失败')

    def test_Function_AcuDC320_Sprint2_002_08_case8(self):
        """
        验证掉电重启，Echilog日志非易失性
        """
        max_num = self.utils.construct_EClig_logs(20, clear=False, connect=False)
        self.utils.reboot_device()
        # 查看当前日志数量是否等于重启前数量
        self.utils.connect_device()
        time.sleep(2)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Echilog')
        current_log_num = int(self.helper.quick_ocr_by_config('EC Used Records'))
        assert current_log_num == max_num, f'重启前后交易日志数量不一致，重启前{max_num}，重启后{current_log_num}'

    def test_Function_AcuDC320_Sprint2_002_08_case1(self):
        """事件日志最大记录6000条验证"""
        self.utils.construct_EClig_logs(6000, clear=False, connect=False)
        with ModbusClient(ModbusProtocol.TCP) as new_client:
            new_get_loss_status = new_client.parse_data(
                new_client.send_custom_message('00 01 00 00 00 09 01 03 10 22 00 01')
            )
            if new_get_loss_status == 1:
                new_status_code = '00'
                new_next_status_code = '01'
            else:
                new_status_code = '01'
                new_next_status_code = '00'
            new_client.send_custom_message(f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {new_status_code}')
            new_client.send_custom_message(f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {new_next_status_code}')
            current_log_num = new_client.parse_data(new_client.validate_register_value('Record Number'))
        assert current_log_num == 6000, "可以写入第6001条交易日志"

    def test_Function_AcuDC320_Sprint2_002_01_case8(self):
        """构建Event Log满，系统状态：Fatal Error，生成一条Echilog，记录旧新值、时间戳、ID"""
        self.utils.construct_EClig_logs(6000, clear=False, connect=False)
        self.utils.connect_device()
        result = self.utils.read_echilog_multi_line(1)
        time_value, type_value, old_value, new_value = result[0]
        pytest.assume(type_value == "Echilog Full",
                      f"类型不正确，期望: Echilog Full, 实际: {type_value}")

    def test_Function_AcuDC320_Sprint2_002_01_case11(self):
        """系统状态：致命错误Fatal error ，提示用户，禁止启动充电交易"""
        self.utils.connect_device()
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')
        # 点击Start_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\operate_failed'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\End_Charging')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            pytest.fail('致命错误Fatal error情况下，仍然可以充电')


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html"])
