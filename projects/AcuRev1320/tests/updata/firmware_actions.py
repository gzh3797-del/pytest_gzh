"""AcuRev-1320 固件升级动作库（updata 用例集共用）。

把 `projects/AcuRev1320/QT_Auto/test_firmware_new.py` 里已跑通的 TCP / RTU 升级
GUI 流程抽成独立函数，供 tests/updata/ 下各用例脚本复用，避免重复维护。

底层仍是「pywinauto 取窗口矩形 + pyautogui 图像/坐标点击」：
  - 图像模板点击（Operation/firmware/Select_All/Connect/Yes/OK 及结果判定图）→ 自定位、可移植
  - 坐标点击（Scan mode 开关、波特率下拉、COM 口行、关闭 Add Connection）→ 经 FirmwareLayout 解析

所有升级结果判定沿用 QT_Auto 的屏幕轮询逻辑：
  Write_Success → 通过；Connect_Failed / Write_Failed → 失败；超时 → 失败。
"""
import os
import sys
import time

# 让 import 在任意调用方式下都成立：仓库根（comm.* / modbus_config）+ QT_Auto（firmware_layout）。
_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(_HERE))))  # updata→tests→AcuRev1320→projects→autotest
_QT_AUTO = os.path.join(_REPO_ROOT, 'projects', 'AcuRev1320', 'QT_Auto')
for _p in (_REPO_ROOT, _QT_AUTO):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from comm.ctl_acuview import dpi  # noqa: E402,F401  DPI 感知须最先设置（在 pyautogui 之前）
import pyautogui  # noqa: E402
import pytest  # noqa: E402
from firmware_layout import FirmwareLayout  # noqa: E402
from modbus_config import modbus_config  # noqa: E402  读 QT_rtu.com_port 等配置

# 升级控件坐标表（复用 QT_Auto 标定结果）
LAYOUT_JSON = os.path.join(_QT_AUTO, 'data', 'firmware_layout.json')

# 升级包路径（与 QT_Auto 共用 data 目录；target = 目标版本，base = 回退基线版本）。
# 当前仓库仅一个 .MFEA，target == base；准备好两个版本后分别指向即可实现往返刷写。
_DATA_DIR = os.path.join(_QT_AUTO, 'data')
PACKAGE_TARGET = os.path.join(_DATA_DIR, 'AcuRev-1320_Application_v1.01p10_20260522.MFEA')
PACKAGE_BASE = os.path.join(_DATA_DIR, 'AcuRev-1320_Application_v1.01p10_20260522.MFEA')

# 公共图像元素前缀
_IMG = r'page_elements\Acuview_public\main_page'

# Select All 禁用(灰)态模板：固件文件解析完成前 Select All 为灰色不可点。
# 绿色可点态模板 Select_All.png 在默认置信度下也会匹配灰色态，单独用它会“点在灰色按钮上”而无效，
# 故以「禁用态消失」作为按钮已转绿可点的判据（见 _wait_select_all_enabled）。
# ⚠️ 文件名必须用 ASCII：cv2.imread（pyautogui confidence 匹配底层）在 Windows 读不了含中文的路径，
#    中文名模板会静默匹配失败 → 等待退化为盲点。故所有「态」模板统一用 ASCII 名。
_SELECT_ALL_DISABLED = rf'{_IMG}\Select_All_Disabled'

# RTU 升级专用「态」模板（需在真机标定抓图，否则对应等待会按“已就绪”放行）：
#   Connect_Disabled  —— 未勾选 COM 口前 Connect 为灰色不可点；以「禁用态消失」判定 Connect 已可点。
#   BaudRate_DropDown —— 波特率下拉真正展开后的可识别标志（展开列表边框/选项行）；
#                        以其「出现」判定下拉已展开，避免下拉没弹出就点选项而点空。
_CONNECT_DISABLED = rf'{_IMG}\Connect_Disabled'
_BAUD_DROPDOWN = rf'{_IMG}\BaudRate_DropDown'

# 非法固件文件加载后上位机弹出的错误提示窗（标题 "Invalid Firmware Data!"）。
# 实机行为：选非法包后设备列表仍会出现「可升级态」（Select All 甚至为灰态也会被 Select_All.png 命中），
# 故不能用「Select All 是否出现」判拒绝；正确信号是这个独立错误弹窗。
_FIRMWARE_INVALID_POPUP = rf'{_IMG}\Invalid_Firmware_Data'


def _wait_select_all_enabled(helper, appear_timeout: int = 15, enable_timeout: int = 180):
    """等待 Select All 由禁用(灰)态变为可点(绿)态后再返回。

    1) 先等禁用态出现，确认已进入「固件文件解析中」；若 appear_timeout 内未捕获，
       视为已解析完成（解析极快）直接放行。
    2) 再等禁用态消失，即按钮转绿可点。enable_timeout 内未消失则判失败。
    """
    if not helper.check_image_exists(_SELECT_ALL_DISABLED, timeout=appear_timeout):
        helper.logger.info('未捕获到 Select All 禁用态，按已可点处理')
        return
    if not helper.check_image_not_exists(_SELECT_ALL_DISABLED, timeout=enable_timeout):
        pytest.fail(f'Select All 在 {enable_timeout}s 内仍为禁用(灰)态，固件文件可能未解析完成')
    helper.logger.info('✅ Select All 已变为可点(绿)态')


def _select_baud(helper, layout, baud_rate, open_timeout: int = 8, retries: int = 3):
    """打开波特率下拉并选中目标值：等下拉「真正展开」再点选项，避免点空。

    根因：原实现 click(BaudRate_ComboBox) 后紧接着 click_baud()，两个坐标点击零间隔，
    下拉尚未弹出/渲染完，第二次点击落在下拉框外 → 波特率没切上。
    这里以 _BAUD_DROPDOWN 展开态「出现」作为可点选项的判据，未展开则重开下拉重试。
    """
    for attempt in range(1, retries + 1):
        layout.click('BaudRate_ComboBox')
        # 置信度降到 0.85：下拉首开时高亮行随「当前选中波特率」移动、顶部值框也随之变，
        # 0.95 整图匹配会因不同用例的高亮/顶值差异而漏判；0.85 仍能稳定命中下拉结构，
        # 且下拉收起时该区域是页面背景、差异大，不会误判。
        if helper.check_image_exists(_BAUD_DROPDOWN, timeout=open_timeout, confidence=0.85):
            layout.click_baud(str(baud_rate))
            if helper.check_image_not_exists(_BAUD_DROPDOWN, timeout=5):  # 选中后下拉收起
                helper.logger.info(f'✅ 波特率已选择 {baud_rate}')
                return True
        helper.logger.warning(f'⚠️ 波特率下拉第 {attempt}/{retries} 次未检测到展开态，重试')
        time.sleep(1)
    # 重试耗尽仍未检测到展开态：多半是 BaudRate_DropDown.png 未标定。
    # 退化为原盲点行为（不比改动前更差），并提示标定模板以恢复稳态等待。
    helper.logger.warning('⚠️ 始终未检测到波特率下拉展开态，退化为盲点点击；'
                          '建议标定 BaudRate_DropDown.png 以启用稳态等待')
    layout.click('BaudRate_ComboBox')
    layout.click_baud(str(baud_rate))
    return False


def _select_com_until_connect_enabled(helper, layout, com_port, scan_timeout: int = 90):
    """轮询勾选指定 COM 口复选框，直到 Connect 由禁用(灰)态转为可点。

    根因：打开 Scan mode 后上位机要扫描串口，扫描期间 COM 行复选框与 Connect 均为灰态；
    原实现只 sleep(2) 就点 COM + Connect，点在灰态上无效 → Connect 没触发 →
    Connect Setting(OK) 对话框不出现 → 后续 click_image(OK) 超时 FAIL。
    这里以 _CONNECT_DISABLED「禁用态消失」作为 COM 已选中、Connect 已可点的判据，
    在扫描超时内反复点 COM 复选框（覆盖扫描未完成时复选框尚不可选的情况）。
    """
    if not helper.check_image_exists(_CONNECT_DISABLED, timeout=5):
        # 未捕获到禁用态：可能模板未标定，或扫描极快已可点。点一次 COM 后直接放行。
        helper.logger.info('未捕获到 Connect 禁用态，按已可点处理（点一次 COM 复选框）')
        layout.click_com(str(com_port))
        return True
    deadline = time.time() + scan_timeout
    while time.time() < deadline:
        layout.click_com(str(com_port))
        if helper.check_image_not_exists(_CONNECT_DISABLED, timeout=4):
            helper.logger.info('✅ COM 已勾选，Connect 已转为可点态')
            return True
        helper.logger.info('Connect 仍为灰态（扫描未完成/复选框未选中），重试勾选 COM 复选框')
        time.sleep(2)
    pytest.fail(f'扫描 {scan_timeout}s 内 Connect 始终为灰态：COM {com_port} 可能未被扫描到或复选框未选中')


def _maximize_main_window(helper):
    """把 Acuview 主窗口最大化，对齐 firmware_layout.json 的基线坐标系（基线＝主窗口最大化）。

    所有 'main' 锚点的相对坐标都按最大化窗口标定；若窗口非最大化（Acuview 记忆了上次的
    非最大化状态），坐标会整体错位，故每次升级前显式最大化，不依赖窗口记忆状态。
    """
    from pywinauto import Desktop

    def _area(w):
        r = w.rectangle()
        return (r.right - r.left) * (r.bottom - r.top)

    try:
        wins = Desktop(backend='uia').windows(title_re='.*Acuview 2.*', top_level_only=True)
        if not wins:
            helper.logger.warning('⚠️ 未找到 Acuview 主窗口，跳过最大化')
            return
        max(wins, key=_area).maximize()
        time.sleep(1)
        helper.logger.info('🗖 主窗口已最大化（对齐基线坐标系）')
    except Exception as exc:  # noqa: BLE001  最大化失败不致命，继续后续流程
        helper.logger.warning(f'⚠️ 最大化主窗口失败（继续）：{exc}')


def make_layout(helper):
    """为给定 helper 构造 FirmwareLayout（坐标解析驱动）。"""
    return FirmwareLayout(helper, LAYOUT_JSON)


def update_firmware_tcp(helper, package_path, device_image_path):
    """TCP 升级流程（启动 Acuview → 连接 → 选包 → Select All → Connect → 等结果）。"""
    helper.hotkey('win', 'd')
    helper.wait(1)
    helper.launch_app(_app_path())
    helper.connect_device(device_image_path)
    time.sleep(5)
    _maximize_main_window(helper)  # 对齐基线坐标系（坐标按主窗口最大化标定）

    # 连上表后进入升级页面较慢，Operation/firmware 给足超时，避免页面未就绪时点空。
    helper.click_image(rf'{_IMG}\Operation', timeout=20)
    helper.click_image(rf'{_IMG}\firmware', timeout=15)
    helper.click_image(rf'{_IMG}\Select_Firmware_File')
    helper.paste_text(package_path)
    pyautogui.hotkey('enter')

    # 等固件文件解析完成：Select All 由禁用(灰)态转为可点(绿)态后再点，避免点在灰色按钮上无效。
    _wait_select_all_enabled(helper)
    helper.click_image(rf'{_IMG}\Select_All', timeout=15)
    helper.click_image(rf'{_IMG}\Connect')
    helper.click_image(rf'{_IMG}\Yes', timeout=10)
    helper.click_image(rf'{_IMG}\Yes', timeout=10)

    return _wait_for_update_completion(helper)


def update_firmware_rtu(helper, layout, package_path, baud_rate):
    """RTU 升级流程（Scan mode + 选波特率 + 选 COM 口 + 升级）。

    前置：firmware_layout.json 的 AddConn_Close 须先在真机标定（否则会 fail-fast）。
    """
    helper.hotkey('win', 'd')
    helper.wait(1)
    helper.launch_app(_app_path())
    layout.click('AddConn_Close')  # 关闭 Add Connection 界面
    _maximize_main_window(helper)  # 对齐基线坐标系（坐标按主窗口最大化标定）

    # 进入升级页面较慢，Operation/firmware 给足超时，避免页面未就绪时点空。
    helper.click_image(rf'{_IMG}\Operation', timeout=20)
    helper.click_image(rf'{_IMG}\firmware', timeout=15)

    helper.click_image(rf'{_IMG}\Yes')  # 点掉扫描弹框
    layout.click('ScanMode_Toggle')     # 打开 Scan mode
    helper.click_image(rf'{_IMG}\Yes')

    # 选择波特率：等下拉真正展开再点选项（原盲点会点空，导致波特率没切上）。
    _select_baud(helper, layout, baud_rate)

    helper.click_image(rf'{_IMG}\Select_Firmware_File')  # 选择升级包
    helper.paste_text(package_path)
    pyautogui.hotkey('enter')
    time.sleep(2)

    # 选中配置指定 COM 口（设备所在口，如 Com 11）的 Select 复选框，而非固定第一行。
    com_port = modbus_config.get('QT_rtu', {}).get('com_port')
    if com_port is None:
        pytest.fail('configs/global.yaml 未配置 QT_rtu.com_port（RTU 升级设备所在 COM 口）')
    # 等扫描完成、COM 复选框与 Connect 转为可点（非灰态）后再点，避免点在灰态上无效。
    _select_com_until_connect_enabled(helper, layout, com_port)

    helper.click_image(rf'{_IMG}\Connect')
    helper.click_image(rf'{_IMG}\OK')   # Connect Setting 对话框保持默认参数，直接 OK
    helper.click_image(rf'{_IMG}\Yes')  # OK 后的确认弹窗

    return _wait_for_update_completion(helper)


def expect_invalid_firmware_file(helper, device_image_path, invalid_file_path):
    """加载一个非 Accuenergy 加密签名的 .MFEA，期望上位机弹出「Invalid Firmware Data!」错误窗。

    用于 case10 的反向校验：返回 True 表示出现了该错误弹窗（文件被正确拒绝）。
    """
    helper.hotkey('win', 'd')
    helper.wait(1)
    helper.launch_app(_app_path())
    helper.connect_device(device_image_path)
    time.sleep(5)
    _maximize_main_window(helper)  # 对齐基线坐标系（坐标按主窗口最大化标定）

    helper.click_image(rf'{_IMG}\Operation', timeout=20)
    helper.click_image(rf'{_IMG}\firmware', timeout=15)
    helper.click_image(rf'{_IMG}\Select_Firmware_File')
    helper.paste_text(invalid_file_path)
    pyautogui.hotkey('enter')

    # 非法文件：上位机解析后弹出「Invalid Firmware Data!」错误窗（即便设备列表仍出现可升级态）。
    # 以该错误弹窗出现作为「被正确拒绝」的判据；轮询等待解析+弹窗。
    return helper.check_image_exists(_FIRMWARE_INVALID_POPUP, timeout=15)


def _wait_for_update_completion(helper, timeout=1800):
    """等待升级完成（30 分钟超时，每 10s 轮询屏幕）。成功返回 True，否则 pytest.fail。"""
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(10)
        if helper.check_image_exists(rf'{_IMG}\Connect_Failed'):
            pytest.fail('升级失败：连接失败（Connect_Failed）')
        elif helper.check_image_exists(rf'{_IMG}\Write_Failed'):
            pytest.fail('升级失败：写入失败（Write_Failed）')
        elif helper.check_image_exists(rf'{_IMG}\Write_Success'):
            helper.logger.info('✅ 升级完成（Write_Success）')
            return True
        helper.keep_active()  # 防锁屏（轻移鼠标，不点击）
    pytest.fail('升级超时：未在规定时间内读到升级结果')


def _app_path():
    return modbus_config['QT_path']
