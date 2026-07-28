r"""Function_AC meter_009_03_case6
用例标题: 上位机setting的Current&Wiring界面，CT Primary越界值写入被拒绝
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Current&Wiring 页 Channel A CT Primary 逐个尝试 4(低于下限5)、2001(高于上限2000)、
          -1(负)、abc(非数字)、空值, 每次尝试保存后经 Modbus 回读 CT Primary(0x104A=4170)。
预期结果: 全部越界/非法值被拒(上位机报错, MANUAL 目视), Modbus 回读保持原值不变; 还原=1000。

生成说明: run_reject_case(非法输入拒绝类)。CT Primary 为 lineEdit(键入, 不依赖 OCR)。判据=
回读!=非法值。CT Primary 合法范围依用例表 5-2000(地址表原文 200/400/800/1200 为标准型档位,
与固件连续可配存在待确认差异, 已在项目 README 固件问题包记录)。
⚠️ 依赖 COM11(RS485) 可用 + 桌面未锁。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.gui_driver import is_session_locked
from comm.ctl_acuview.testcase_engine import run_reject_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面")


def test_009_03_case6():
    report = run_reject_case(
        case_meta={
            '编号': 'Function_AC meter_009_03_case6',
            '标题': '上位机setting的Current&Wiring界面，CT Primary越界值写入被拒绝',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
            '测试步骤': 'Channel A CT Primary 尝试 4/2001/-1/abc/空, 每次 Modbus 回读 0x104A 确认未变',
            '预期结果': '越界/非法值被拒, Modbus 回读保持原值; 还原=1000',
        },
        register=4170, page="Current_Wire", widget="Channel_A_CT_Primary_LineEdit",
        illegal_values=[4, 2001, -1, "abc", ""], restore_value=1000, config_path=TEST_CONFIG,
    )
    assert report.passed
