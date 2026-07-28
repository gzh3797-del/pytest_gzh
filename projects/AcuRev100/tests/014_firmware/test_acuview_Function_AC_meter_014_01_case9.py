r"""Function_AC meter_014_01_case9
用例标题: 上位机通过RTU方式，压力升级15次
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: RTU连接; 连续升级压测15次(标题15次/步骤写10次, 按标题从严取15)
预期结果: 每次均升级成功, 过程无错误

生成说明: run_firmware_update_case(rounds=15), 波特率保持当前值。每轮判 Write Success, 轮间等待10s(表重启缓冲)。判据(2026-07-15 裁决): 同版重刷判"升级流程成功", FW版本记录不强判(firmware.expect_version 填值后自动严判); 数据保持用配置类寄存器快照。🔴 升级会重启电表, 授权门禁 skip 兜底。
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


def test_014_01_case9():
    report = run_firmware_update_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case9',
            '标题': '上位机通过RTU方式，压力升级15次',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'RTU连接; 连续升级压测15次(标题15次/步骤写10次, 按标题从严取15)',
            '预期结果': '每次均升级成功, 过程无错误',
        },
        config_path=TEST_CONFIG, rounds=15,
    )
    assert report.passed
