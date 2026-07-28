r"""Function_AC meter_015_01_case1
用例标题: 上位机上通过RTU,断开输入源，重置能量 (clear energy),重置成功
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 断开输入源(电流置0)-> Reading->Energy 页点 Clear Energy 按钮 + 确认弹窗;
          经 Modbus 回读各相/系统能量=0。
预期结果: 上位机页面能量被清 0(Modbus 回读 System/Phase Active Energy Import 等 = 0)。

实现: 先脚本控源(CL3021)把三相电流置 0(保留 Ua 供电, 忠实"断开输入源"), 使能量停止累计;
  再 run_button_action_case 点 Energy_Clear_Button + _confirm_dialogs 确认, Modbus 严格回读能量=0
  (无电流 → 清零后不再累计); 最后恢复源保活。
🔴红线-能量清零: 2026-07-15 用户在场授权放行。ADC 已恢复正常, 能量真在累计, 故必须先断源电流
  才能严格验 0(否则清后瞬间又微量再累计)。
"""
import time
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_button_action_case
from projects.AcuRev100.tests.helpers_accuracy import (
    ensure_source_keepalive, exit_idle)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_015_01_case1():
    # 1) 断开输入源: 三相电流置 0(保留 Ua 供电), 能量停止累计
    src = ensure_source_keepalive(TEST_CONFIG)
    time.sleep(3)   # 电流归零后留时, 确保能量不再增长
    try:
        # 2) GUI 清能量 + Modbus 严格回读=0(无电流, 清后不再累计)
        report = run_button_action_case(
            case_meta={
                '编号': 'Function_AC meter_015_01_case1',
                '标题': '上位机上通过RTU,断开输入源，重置能量 (clear energy),重置成功',
                '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
                '测试步骤': '断开输入源(电流0)-> Energy 页点 Clear Energy + 确认; Modbus 回读能量=0',
                '预期结果': '各相/系统能量被清 0',
            },
            page="Energy", button_widget="Energy_Clear_Button",
            verify=[
                ("清能量后 系统有功输入电能=0", 18950, 0),
                ("清能量后 PhaseA 有功输入电能=0", 18944, 0),
            ],
            config_path=TEST_CONFIG,
        )
    finally:
        # 3) 恢复源保活(退场)
        exit_idle(src, TEST_CONFIG)
    assert report.passed
