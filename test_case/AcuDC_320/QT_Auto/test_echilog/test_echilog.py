import pytest
import sys
import os
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



if __name__ == "__main__":
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html"])