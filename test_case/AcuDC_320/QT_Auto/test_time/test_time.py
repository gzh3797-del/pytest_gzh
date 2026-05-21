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


class TestTime:
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

    @pytest.mark.parametrize("time", [
        '20240819000000',
        '20240823230000',
        '20240824235900',
        '20240825235859',
        '20240229235800'
    ])
    def test_Function_AcuDC320_01_01_case1(self, time):
        """验证设备实时时钟功能"""
        self.utils.connect_device()
        self.utils.set_custom_time(time)

    def test_Function_AcuDC320_01_02_case3(self):
        """铅封封闭，可通过Modbus TCP协议设置时间"""
        with ModbusClient(ModbusProtocol.RTU) as new_client:
            res1 = new_client.send_custom_message(
                '01 10 10 3F 00 07 0E 00 05 07 E8 00 03 00 01 00 00 00 17 00 14  ')
        pytest.assume(res1[3:5] == '10', f"时间写入失败")
