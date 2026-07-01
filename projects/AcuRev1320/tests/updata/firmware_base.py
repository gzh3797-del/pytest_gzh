"""updata 固件升级用例公共基类：负责 helper / layout 构造与安全 teardown。"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(_HERE))))
_QT_AUTO = os.path.join(_REPO_ROOT, 'projects', 'AcuRev1320', 'QT_Auto')
for _p in (_REPO_ROOT, _QT_AUTO, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from comm.ctl_acuview import dpi  # noqa: E402,F401  DPI 感知须最先设置
import pytest  # noqa: E402
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper  # noqa: E402
from modbus_config import modbus_config  # noqa: E402
import firmware_actions as fa  # noqa: E402


class FirmwareTestBase:
    """各升级用例类继承本基类，复用同一套 setup / teardown。"""

    @pytest.fixture(autouse=True)
    def setup(self, request):
        # manual 用例只在 body 里 pytest.skip 记录手动步骤，无需启动 GUI / 杀 Acuview。
        if request.node.get_closest_marker('manual') is not None:
            self.helper = None
            yield
            return
        self.app_path = modbus_config['QT_path']
        self.device_image_path = modbus_config['device_image_path']
        self.helper = AutoHelper(confidence=0.8, click_pause=0.4, image_pause=0.4)
        self.layout = fa.make_layout(self.helper)
        self.test_name = request.node.name
        # 升级已开始刷写、结果尚未确认时为 True；此时禁止关 Acuview，避免中途强杀设备。
        self._flashing = False
        self.helper.kill_acuview_apps()
        yield
        if self._flashing:
            self.helper.logger.warning(
                '⚠️ 升级状态未确认（疑似升级中途异常），已跳过关闭 Acuview，保留窗口供人工检查')
        else:
            self.helper.kill_acuview_apps()

    # ---- 包装升级动作：进入即标记 _flashing，Write_Success / 失败后由动作库判定 ----
    def _do_tcp(self, package_path):
        self._flashing = True
        ok = fa.update_firmware_tcp(self.helper, package_path, self.device_image_path)
        self._flashing = False
        return ok

    def _do_rtu(self, package_path, baud_rate):
        self._flashing = True
        ok = fa.update_firmware_rtu(self.helper, self.layout, package_path, baud_rate)
        self._flashing = False
        return ok
