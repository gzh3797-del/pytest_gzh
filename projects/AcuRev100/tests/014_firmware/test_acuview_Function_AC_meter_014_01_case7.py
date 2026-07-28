r"""Function_AC meter_014_01_case7
用例标题: 其他产品升级包进行升级（固件包合法性校验）
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 升级界面选择他产品升级包(4100)加载
预期结果: 升级文件解析失败, 提示'固件文件无效'(Invalid Firmware Data!)

生成说明: run_firmware_invalid_file_case: 选 AcuRev-4100 .MFEA, 断言报错弹窗出现。不刷机不重启, 无需刷机授权门禁。非法包路径见 config firmware.invalid_package。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.firmware_update import run_firmware_invalid_file_case
from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
]


def test_014_01_case7():
    report = run_firmware_invalid_file_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case7',
            '标题': '其他产品升级包进行升级（固件包合法性校验）',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': '升级界面选择他产品升级包(4100)加载',
            '预期结果': '升级文件解析失败, 提示"固件文件无效"(Invalid Firmware Data!)',
        },
        config_path=TEST_CONFIG,
    )
    assert report.passed
