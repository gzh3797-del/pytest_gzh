r"""Function_AC meter_014_01_case16
用例标题: 升级前后，查看电表各个寄存器数据不会发生变化
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 升级前拍配置寄存器+设备信息快照; 升级; 升级后重拍比对
预期结果: 升级前后寄存器数据一致(需量/datalog类动态数据偏差注意识别)

生成说明: run_firmware_update_case(check_info_keep=True): Basic Setting 101项快照 + SN/Model/HW/Boot版本保持比对。能量/实时测量为动态量不作等值判据(2026-07-15 裁决)。判据(2026-07-15 裁决): 同版重刷判"升级流程成功", FW版本记录不强判(firmware.expect_version 填值后自动严判); 数据保持用配置类寄存器快照。🔴 升级会重启电表, 授权门禁 skip 兜底。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.firmware_update import run_firmware_update_case, upgrade_allowed
from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
    pytest.mark.skipif(not upgrade_allowed(TEST_CONFIG),
                       reason="刷机未授权(run.allow_firmware_upgrade=false); 升级会重启电表, 现场确认后开启"),
]


def test_014_01_case16():
    report = run_firmware_update_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case16',
            '标题': '升级前后，查看电表各个寄存器数据不会发生变化',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': '升级前拍配置寄存器+设备信息快照; 升级; 升级后重拍比对',
            '预期结果': '升级前后寄存器数据一致(需量/datalog类动态数据偏差注意识别)',
        },
        config_path=TEST_CONFIG, check_info_keep=True,
    )
    assert report.passed
