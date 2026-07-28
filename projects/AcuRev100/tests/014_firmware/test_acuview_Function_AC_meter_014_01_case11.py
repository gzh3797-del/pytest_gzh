r"""Function_AC meter_014_01_case11
用例标题: 上位机升级后，接入源，查看电压、电流、角度精度
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: RTU连接; 升级; 升级后接源输入电压/电流/角度查精度
预期结果: 升级成功; 电压、电流和角度符合精度要求

生成说明: run_firmware_update_case(升级部分自动)。精度部分=升级后复跑 002/004/006 精度批(接源判据/救源机制齐全), 本用例记 MANUAL 指引, 不重复造精度判据。判据(2026-07-15 裁决): 同版重刷判"升级流程成功", FW版本记录不强判(firmware.expect_version 填值后自动严判); 数据保持用配置类寄存器快照。🔴 升级会重启电表, 授权门禁 skip 兜底。
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


def test_014_01_case11():
    report = run_firmware_update_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case11',
            '标题': '上位机升级后，接入源，查看电压、电流、角度精度',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'RTU连接; 升级; 升级后接源输入电压/电流/角度查精度',
            '预期结果': '升级成功; 电压、电流和角度符合精度要求',
        },
        config_path=TEST_CONFIG, physical_note="接源查电压/电流/角度精度: 升级完成后复跑精度批 pytest projects/AcuRev100/tests/002_phase_voltage 004_current 006_phase_angle",
    )
    assert report.passed
