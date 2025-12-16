import pyautogui
import pytest
import time
import sys
import os
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from modbus_config import modbus_config
'''
环境准备:
1、config文件配置， 例如:
  "QT_path": "C:\\Users\\ZanHu\\Acuview2\\Acuview 2.exe",
  "device_image_path": "page_elements\\AucRev1320\\1320 TCP",
  "QT_tcp": {
   "host": "192.168.2.121",
   "port": 502,
    "timeout": 10,
    "slave_id": 1
2、 新建TCP连接session， TCP连接命名称为1320 TCP, 确保上位机手动可以连接到电表
3、 设备TCP连接的截图元素存放在工程page_elements\\AucRev1320\\目录下，命名1320 TCP，格式png
4、 设置升级包路径
'''
# 升级包路径
package_path_target = r'C:\autotest_local\update_version\AcuRev-4100_Application_v1.01p35_20251126.MFEA'
package_path_base = r'C:\autotest_local\update_version\AcuRev-4100_Application_v1.01p32_20251025.MFEA'

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestAcuviewAutomation:
    """Acuview 自动化测试类"""

    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        self.helper.kill_acuview_apps()
        yield
        self.helper.kill_acuview_apps()

    # ---------------- 核心升级方法 ----------------
    def update_firmware_tcp(self, package_path):
        """TCP 升级流程"""
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        self.helper.connect_device(self.device_image_path)
        time.sleep(10)

        self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
        self.helper.paste_text(package_path)
        pyautogui.hotkey('enter')
        time.sleep(5)

        # 勾选TCP坐标
        self.helper.click_pos((623, 402))
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')

        self._wait_for_update_completion()

    def update_firmware_rtu(self, package_path, baud_rate_pos):
        """RTU 升级流程"""
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        self.helper.click_pos((1410, 90))  # 关闭 add_connection 界面

        self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')

        # 点掉扫描弹框
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
        # 打开Scan mode,笔记本上坐标609，282
        self.helper.click_pos((609, 282))
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')

        # 选择波特率，笔记本上坐标1110，325
        self.helper.click_pos((1110, 325))
        self.helper.click_pos(baud_rate_pos)

        # 选择升级包
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
        self.helper.paste_text(package_path)
        pyautogui.hotkey('enter')
        time.sleep(2)

        # 选中第一个com口
        self.helper.click_pos((624, 405))
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\OK')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')

        self._wait_for_update_completion()

    def _wait_for_update_completion(self, timeout=1800):
        """等待升级完成"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            time.sleep(20)
            if self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Connect_Failed'):
                pytest.fail('升级失败：连接失败')
            elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Failed'):
                pytest.fail('升级失败：写入失败')
            elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Success'):
                self.helper.logger.info('✅ 升级完成')
                return
            # 防止屏幕锁屏
            self.helper.click_pos((1269, 500))
        pytest.fail('升级超时')

    # ---------------- 动态生成 test 方法 ----------------
    def generate_test_method_tcp_update(self):
        """生成 TCP 升级测试方法"""
        def test_method(self):
            self.update_firmware_tcp(package_path_target)
            self.update_firmware_tcp(package_path_base)
        return test_method

    def generate_test_method_rtu_update(self, baud_rate_pos):
        """生成 RTU 升级测试方法"""
        def test_method(self):
            self.update_firmware_rtu(package_path_target, baud_rate_pos)
            self.update_firmware_rtu(package_path_base, baud_rate_pos)
        return test_method

    # ---------------- 注册动态测试 ----------------
    @classmethod
    def generate_tests_tcp(cls, rounds=2):
        """动态生成 TCP 升级测试方法"""
        for i in range(rounds):
            test_method = cls.generate_test_method_tcp_update(cls)
            test_method.__name__ = f"test_TCP_update_round{i}"
            setattr(cls, test_method.__name__, test_method)

    @classmethod
    def generate_tests_rtu(cls, rounds=2):
        """动态生成 RTU 升级测试方法"""
        pos_dic = {
            '115200': (1077, 416),
            '57600': (1077, 399),
            '38400': (1077, 382),
            '19200': (1077, 365),
            '9600': (1077, 348)
        }
        for j in range(rounds):
            for rate, pos in pos_dic.items():
                test_method = cls.generate_test_method_rtu_update(cls, pos)
                test_method.__name__ = f"test_RTU_update_{rate}_round{j+1}"
                setattr(cls, test_method.__name__, test_method)


# ------------------- 生成测试 -------------------
TestAcuviewAutomation.generate_tests_tcp(rounds=2)
# TestAcuviewAutomation.generate_tests_rtu(rounds=1)

# ------------------- 直接运行 -------------------
if __name__ == "__main__":
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html", '-x'])
