r"""Function_AC meter_014_01_case2
用例标题: 上位机通过RTU方式，波特率设置为19200,升级正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: RTU连接; 记FW/配置基线; 波特率19200升级; 查无错/版本/数据保持(测量数据/日志不变,新版本继续工作)
预期结果: 升级过程无错误; FW版本正确; 配置数据保持

生成说明: run_firmware_update_case(baud=19200)。2026-07-15 晚该波特率已真机手动走通一轮(Write Success)。判据(2026-07-15 裁决): 同版重刷判"升级流程成功", FW版本记录不强判(firmware.expect_version 填值后自动严判); 数据保持用配置类寄存器快照。🔴 升级会重启电表, 授权门禁 skip 兜底。
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


def test_014_01_case2():
    report = run_firmware_update_case(
        case_meta={
            '编号': 'Function_AC meter_014_01_case2',
            '标题': '上位机通过RTU方式，波特率设置为19200,升级正常',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'RTU连接; 记FW/配置基线; 波特率19200升级; 查无错/版本/数据保持(测量数据/日志不变,新版本继续工作)',
            '预期结果': '升级过程无错误; FW版本正确; 配置数据保持',
        },
        config_path=TEST_CONFIG, baud=19200,
        physical_note="升级后测量准确性由 case11+精度批覆盖",
    )
    assert report.passed
