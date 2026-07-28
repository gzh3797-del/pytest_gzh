r"""Function_AC meter_014_01_case1
用例标题: 上位机通过RTU方式，波特率设置为9600,升级正常。
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: RTU连接→记FW版本与配置基线→选升级包(9600)→升级→查过程无错→查FW版本→查数据保持
预期结果: 升级过程无错误; FW版本正确; 配置数据保持; 测量数据准确(电压,电流,功率)

生成说明: run_firmware_update_case(baud=9600)。流程/坐标按 2026-07-15 真机实操记录。
判据(2026-07-15 裁决): 仅一版升级包→同版重刷判"升级流程成功"(Write Success+Update Finished+
恢复在线), FW 版本记录不强判(config firmware.expect_version 填值后自动严判); 数据保持用
配置类寄存器快照(能量持续累计不作保持判据); 测量准确性记 MANUAL(由 case11+精度批覆盖)。
🔴 升级会重启电表: config run.allow_firmware_upgrade=true 且工程师在场才实跑, 否则 skip。
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


def test_014_01_case1():
    report = run_firmware_update_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case1',
            '标题': '上位机通过RTU方式，波特率设置为9600,升级正常',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'RTU连接; 记FW/配置基线; 波特率9600升级; 查无错/版本/数据保持',
            '预期结果': '升级过程无错误; FW版本正确; 配置数据保持',
        },
        config_path=TEST_CONFIG, baud=9600,
        physical_note="升级后测量数据准确性(电压/电流/功率)由 case11+精度批(002/004/005)覆盖",
    )
    assert report.passed
