import pyautogui
import pytest
import time
import sys
import os
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from modbus_config import modbus_config

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestAcuviewAutomation:
    """Acuview自动化测试类"""

    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        # 初始化工具类
        self.helper.kill_acuview_apps()
        self.helper.hotkey('win', 'd')
        self.helper.launch_app(self.app_path)
        yield
        self.helper.kill_acuview_apps()

    def generate_test_method_RTU_update(self, pos):

        def test_method(self):
            self.helper.hotkey('win', 'd')
            self.helper.wait(1)
            self.helper.launch_app(self.app_path)

            self.helper.click_image(r'page_elements\Acuview_public\Add_Connect_page\Add_Connection', offset_x=848,
                                    confidence=0.9)
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
            self.helper.click_pos((1123, 320))
            self.helper.click_pos(pos)
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
            # 粘贴路径
            self.helper.paste_text(self.package_path)
            # 按回车
            pyautogui.hotkey('enter')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Scan_Mode'):
                time.sleep(2)
                self.helper.click_pos((623, 402))
                self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
                self.helper.click_image(r'page_elements\Acuview_public\main_page\OK')
                self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
                time.sleep(10)
            time_out = 1800
            start_time = time.time()
            while time.time() - start_time < time_out:
                time.sleep(20)
                if self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Connect_Failed'):
                    pytest.fail('升级失败：连接失败')
                elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Failed'):
                    pytest.fail('升级失败：写入失败')
                elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Success'):
                    self.helper.logger.info('✅ 升级完成')
                    break
                self.helper.click_pos((1269, 500))

        return test_method

    def generate_test_method_TCP_update(self):

        def test_method(self):
            # 启动应用
            self.helper.hotkey('win', 'd')
            self.helper.wait(1)
            self.helper.launch_app(self.app_path)
            self.helper.connect_device(self.device_image_path)
            time.sleep(5)
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
            # 粘贴路径
            self.helper.paste_text(self.package_path)
            # 按回车
            pyautogui.hotkey('enter')
            time.sleep(2)
            self.helper.click_pos((623, 402))
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
            self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
            time_out = 1800
            start_time = time.time()
            while time.time() - start_time < time_out:
                time.sleep(20)
                if self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Connect_Failed'):
                    pytest.fail('升级失败：连接失败')
                elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Failed'):
                    pytest.fail('升级失败：写入失败')
                elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Success'):
                    self.helper.logger.info('✅ 升级完成')
                    break
                self.helper.click_pos((1269, 500))

        return test_method

    @classmethod
    def generate_tests(cls):
        """动态生成测试方法"""
        instance = cls()
        case_ids = ['Function_AcuDC320_04_01_case1',
                    'Function_AcuDC320_04_01_case2',
                    'Function_AcuDC320_04_01_case3',
                    'Function_AcuDC320_04_01_case4',
                    'Function_AcuDC320_04_01_case5']
        pos_dic = {'115200': (1077, 410), '57600': (1077, 399), '38400': (1077, 382), '19200': (1077, 360),
                   '9600': (1077, 348)}
        rates = ['115200', '57600', '38400', '19200', '9600']
        for case_id in case_ids:
            for rate in rates:
                test_method = instance.generate_test_method_RTU_update(pos_dic[rate])
                test_method.__name__ = f"test_{case_id}_{rate}"
                setattr(cls, test_method.__name__, test_method)
        for j in range(4):
            for rate in rates:
                test_method = instance.generate_test_method_RTU_update(pos_dic[rate])
                test_method.__name__ = f"test_Function_AcuDC320_04_01_case11_{rate}_round{j + 1}"
                setattr(cls, test_method.__name__, test_method)


TestAcuviewAutomation.generate_tests()
if __name__ == "__main__":
    # 直接运行测试
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html"])
