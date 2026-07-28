r"""Function_AC meter_014_01_case10
用例标题: 上位机升级界面显示准确，设备信息栏中Model、Hardware、Firmware等信息显示正确
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: RTU连接; 打开 Firmware Update 界面; OCR 信息栏与 Modbus 寄存器真值比对
预期结果: Model/Hardware/Firmware 显示正确(Model 正式串待固件确认, 当前=MAC1)

生成说明: run_firmware_info_display_case: 信息栏 OCR ↔ MODEL@61553/HARDWARE@61520/FIRMWARE@61440(ASCII解码)比对; Model 期望取 config firmware.expect_model(当前 MAC1, 正式版改 AcuRev-101-mA/mV)。只读不刷机, 无需刷机授权。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.firmware_update import run_firmware_info_display_case
from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
]


def test_014_01_case10():
    report = run_firmware_info_display_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case10',
            '标题': '上位机升级界面显示准确，设备信息栏中Model、Hardware、Firmware等信息显示正确',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'RTU连接; 打开 Firmware Update 界面; OCR 信息栏与 Modbus 寄存器真值比对',
            '预期结果': 'Model/Hardware/Firmware 显示正确(Model 正式串待固件确认, 当前=MAC1)',
        },
        config_path=TEST_CONFIG,
    )
    assert report.passed
