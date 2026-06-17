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


class TestRTUCommunication:
    #
    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.app_root_path = os.path.dirname(self.app_path)
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        # 初始化工具类
        self.utils = CommonUtils(self.helper, self.app_path, self.device_image_path)
        self.helper.kill_acuview_apps()
        self.helper.wait(2)
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        yield
        self.helper.kill_acuview_apps()

    @pytest.mark.parametrize("Parity", [
        'N',
        'O',
        'E',
    ])
    def test_Function_AcuDC320_03_01_case1(self, Parity):
        """通过RS485配置校验位功能正常"""
        dict = {
            'E': 'D46D6DF4DE8A',
            'O': 'D46D6DF4DE8C',
            'N': 'D46D6DF4DE8D'
        }
        self.utils.connect_device()
        self.utils.configure_Parity(Parity)
        self.helper.kill_acuview_apps()
        with ModbusClient(ModbusProtocol.RTU, parity=Parity) as client:
            client.send_custom_message('01 10 10 01 00 01 02 4B 00')
        with ModbusClient(ModbusProtocol.RTU, parity=Parity) as client:
            client.validate_register_value('MAC address', dict[Parity], True)
            mac = client.parse_data(client.validate_register_value('MAC address'), 'mac')
        assert dict[Parity] == mac, 'mac更改失败'

    def test_restore_Parity(self):
        """通过RS485配置校验位功能正常"""
        self.utils.connect_device()
        self.utils.configure_Parity('N')


    @pytest.mark.parametrize("Baud_Rate", [
        2400,
        4800,
        9600,
        19200,
        38400,
        57600,
        78600,
        152000
    ])
    def test_Function_AcuDC320_03_01_case5(self, Baud_Rate):
        """通过RS485配置波特率功能正常"""
        dict = {
            2400: 'D46D6DF4DE81',
            4800: 'D46D6DF4DE82',
            9600: 'D46D6DF4DE83',
            19200: 'D46D6DF4DE84',
            38400: 'D46D6DF4DE85',
            57600: 'D46D6DF4DE86',
            78600: 'D46D6DF4DE87',
            152000: 'D46D6DF4DE88',
        }
        self.utils.connect_device()
        self.utils.configure_Baud_Rate(Baud_Rate)
        self.helper.kill_acuview_apps()
        with ModbusClient(ModbusProtocol.RTU, baudrate=Baud_Rate) as client:
            client.validate_register_value('MAC address', dict[Baud_Rate], True)
            mac = client.parse_data(client.validate_register_value('MAC address'), 'mac')
        assert dict[Baud_Rate] == mac, 'mac更改失败'
    def test_restore_Baud_Rate(self):
        """通过RS485配置校验位功能正常"""
        self.utils.connect_device()
        self.utils.configure_Baud_Rate(19200)
