import os
import sys

# 让 import 在任意调用方式下都成立:
#   - 仓库根(autotest/) → 供 `comm.*` / `modbus_config` 导入
#   - 本目录(QT_Auto/)  → 供 `firmware_layout` 导入
# 必须在导入 comm.* 之前;本目录有独立 pytest.ini,pytest 的 rootdir 落在 QT_Auto,
# 不会加载仓库根 conftest,故须自行把仓库根塞进 sys.path。
_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(_HERE)))  # QT_Auto→AcuRev1320→projects→autotest
for _p in (_REPO_ROOT, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from comm.ctl_acuview import dpi  # noqa: E402,F401  DPI 感知须最先设置(在 pyautogui 之前)
import pyautogui
import pytest
import time
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from firmware_layout import FirmwareLayout
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
package_path_target = r'C:\Users\ZihanGao\Desktop\test_code\autotest\projects\AcuRev1320\QT_Auto\data\AcuRev-1320_Application_v1.01p10_20260522.MFEA'
package_path_base = r'C:\Users\ZihanGao\Desktop\test_code\autotest\projects\AcuRev1320\QT_Auto\data\AcuRev-1320_Application_v1.01p10_20260522.MFEA'

# 固件升级控件坐标表(相对锚点的相对坐标，解析见 firmware_layout.py)
LAYOUT_JSON = os.path.join(_HERE, 'data', 'firmware_layout.json')

# RTU 升级「态」模板(需真机标定；未标定时对应等待按"已就绪"放行，退化为盲点行为)
#   Connect_Disabled  —— 未勾选 COM 口前 Connect 为灰色不可点；以「禁用态消失」判定已可点
#   BaudRate_DropDown —— 波特率下拉展开后的可识别标志；以其「出现」判定下拉已弹出再点选项
# ⚠️ 文件名必须 ASCII：cv2.imread 在 Windows 读不了中文路径，中文名模板会静默匹配失败。
_CONNECT_DISABLED = r'page_elements\Acuview_public\main_page\Connect_Disabled'
_BAUD_DROPDOWN = r'page_elements\Acuview_public\main_page\BaudRate_DropDown'


class TestAcuviewAutomation:
    """Acuview 自动化测试类"""

    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8, click_pause=0.4, image_pause=0.4)
        self.layout = FirmwareLayout(self.helper, LAYOUT_JSON)
        self.test_name = request.node.name
        # 升级是否处于"已开始刷写、但结果尚未确认"的状态；True 时禁止关 Acuview，避免中途强杀设备
        self._flashing = False
        self.helper.kill_acuview_apps()
        yield
        if self._flashing:
            self.helper.logger.warning(
                '⚠️ 升级状态未确认(疑似升级中途异常)，已跳过关闭 Acuview，保留窗口供人工检查升级结果；'
                '确认无误后请手动关闭 Acuview')
        else:
            self.helper.kill_acuview_apps()

    # ---------------- 核心升级方法 ----------------
    def update_firmware_tcp(self, package_path):
        """TCP 升级流程"""
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        self.helper.connect_device(self.device_image_path)
        time.sleep(5)

        self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
        self.helper.paste_text(package_path)
        pyautogui.hotkey('enter')
        time.sleep(5)  # 等固件文件加载完，Select All 才会变为可点(绿色)

        # 选中要升级的设备行(点 Select All；放宽超时，等它变绿可点再点，不依赖坐标)
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_All', timeout=15)
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
        # Connect 后弹"确认升级"对话框可能稍慢，放宽超时避免过早判失败
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes', timeout=10)
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes', timeout=10)

        self._wait_for_update_completion()

    def update_firmware_rtu(self, package_path, baud_rate):
        """RTU 升级流程"""
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)
        self.layout.click('AddConn_Close')  # 关闭 add_connection 界面

        self.helper.click_image(r'page_elements\Acuview_public\main_page\Operation')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\firmware')

        # 点掉扫描弹框
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')
        # 打开Scan mode
        self.layout.click('ScanMode_Toggle')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')

        # 选择波特率(等下拉真正展开再点选项，避免下拉未弹出就点空 → 波特率没切上)
        self._select_baud(baud_rate)

        # 选择升级包
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Select_Firmware_File')
        self.helper.paste_text(package_path)
        pyautogui.hotkey('enter')
        time.sleep(2)

        # 选中第一个 com 口：等扫描完成、COM 复选框与 Connect 转为可点(非灰态)再点，
        # 避免点在灰态上无效 → Connect 没触发 → Connect Setting(OK) 不出现而超时。
        self._select_com_until_connect_enabled()
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Connect')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\OK')
        self.helper.click_image(r'page_elements\Acuview_public\main_page\Yes')

        self._wait_for_update_completion()

    # ---------------- RTU 条件等待(与 tests/updata/firmware_actions.py 同思路) ----------------
    def _select_baud(self, baud_rate, open_timeout=8, retries=3):
        """打开波特率下拉并选中目标值：等下拉「真正展开」再点选项，避免点空。"""
        for attempt in range(1, retries + 1):
            self.layout.click('BaudRate_ComboBox')
            # 置信度降到 0.85：下拉首开时高亮行随「当前选中波特率」移动、顶部值框也随之变，
            # 0.95 整图匹配会因不同用例的高亮/顶值差异而漏判；0.85 仍能稳定命中下拉结构，
            # 且下拉收起时该区域是页面背景、差异大，不会误判。
            if self.helper.check_image_exists(_BAUD_DROPDOWN, timeout=open_timeout, confidence=0.85):
                self.layout.click_baud(baud_rate)
                if self.helper.check_image_not_exists(_BAUD_DROPDOWN, timeout=5):  # 选中后下拉收起
                    self.helper.logger.info(f'✅ 波特率已选择 {baud_rate}')
                    return
            self.helper.logger.warning(f'⚠️ 波特率下拉第 {attempt}/{retries} 次未检测到展开态，重试')
            time.sleep(1)
        # 重试耗尽仍未检测到展开态(多半是模板未标定)：退化为原盲点行为，不比改动前更差。
        self.helper.logger.warning('⚠️ 始终未检测到波特率下拉展开态，退化为盲点点击；'
                                   '建议标定 BaudRate_DropDown.png 以启用稳态等待')
        self.layout.click('BaudRate_ComboBox')
        self.layout.click_baud(baud_rate)

    def _select_com_until_connect_enabled(self, scan_timeout=90):
        """轮询勾选第一个 COM 口复选框，直到 Connect 由禁用(灰)态转为可点。"""
        if not self.helper.check_image_exists(_CONNECT_DISABLED, timeout=5):
            # 未捕获禁用态：模板未标定或扫描极快已可点。点一次 COM 后直接放行。
            self.helper.logger.info('未捕获到 Connect 禁用态，按已可点处理(点一次 COM 复选框)')
            self.layout.click('FirstComPort_Row')
            return
        deadline = time.time() + scan_timeout
        while time.time() < deadline:
            self.layout.click('FirstComPort_Row')
            if self.helper.check_image_not_exists(_CONNECT_DISABLED, timeout=4):
                self.helper.logger.info('✅ COM 已勾选，Connect 已转为可点态')
                return
            self.helper.logger.info('Connect 仍为灰态(扫描未完成/复选框未选中)，重试勾选 COM 复选框')
            time.sleep(2)
        pytest.fail(f'扫描 {scan_timeout}s 内 Connect 始终为灰态：COM 口可能未被扫描到或复选框未选中')

    def _wait_for_update_completion(self, timeout=1800):
        """等待升级完成。进入即标记 _flashing=True(正在刷写)，拿到明确结果后置回 False。"""
        self._flashing = True  # 已开始刷写、结果未定 → teardown 期间禁止关 Acuview
        start_time = time.time()
        while time.time() - start_time < timeout:
            time.sleep(10)
            if self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Connect_Failed'):
                self._flashing = False  # 连接失败，刷写未真正开始，结果已确认
                pytest.fail('升级失败：连接失败')
            elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Failed'):
                self._flashing = False  # 写入失败，刷写已结束，结果已确认
                pytest.fail('升级失败：写入失败')
            elif self.helper.check_image_exists(r'page_elements\Acuview_public\main_page\Write_Success'):
                self._flashing = False  # 升级成功，结果已确认
                self.helper.logger.info('✅ 升级完成')
                return
            # 防止屏幕锁屏(轻移鼠标，不点击，避免误触界面)
            self.helper.keep_active()
        # 超时：升级是否完成无法确认 → 保持 _flashing=True，teardown 将保留窗口供人工检查
        pytest.fail('升级超时：未在规定时间内读到升级结果，保留窗口供人工检查')

    # ---------------- 动态生成 test 方法 ----------------
    def generate_test_method_tcp_update(self):
        """生成 TCP 升级测试方法"""
        def test_method(self):
            self.update_firmware_tcp(package_path_target)
            self.update_firmware_tcp(package_path_base)
        return test_method

    def generate_test_method_rtu_update(self, baud_rate):
        """生成 RTU 升级测试方法"""
        def test_method(self):
            self.update_firmware_rtu(package_path_target, baud_rate)
            self.update_firmware_rtu(package_path_base, baud_rate)
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
        """动态生成 RTU 升级测试方法(波特率坐标查 firmware_layout.json 的 baud_options)"""
        baud_rates = ['115200', '57600', '38400', '19200', '9600']
        for j in range(rounds):
            for rate in baud_rates:
                test_method = cls.generate_test_method_rtu_update(cls, rate)
                test_method.__name__ = f"test_RTU_update_{rate}_round{j+1}"
                setattr(cls, test_method.__name__, test_method)


# ------------------- 生成测试 -------------------
TestAcuviewAutomation.generate_tests_tcp(rounds=2)
# TestAcuviewAutomation.generate_tests_rtu(rounds=1)

# ------------------- 直接运行 -------------------
if __name__ == "__main__":
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html", '-x'])
