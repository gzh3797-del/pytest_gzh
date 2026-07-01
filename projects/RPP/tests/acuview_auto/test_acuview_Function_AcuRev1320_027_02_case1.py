"""自动化用例 Function_AcuRev1320_027_02_case1（密码及背光设置）

由 manual_testcase/manual_test_tmp.xlsx 经 gen-testcase Skill 转换而来。
文件名 = test_acuview_<用例编号>.py(一条用例一个文件)。
判据: Acuview2 界面把 LCD Backlight Time(0x1018) 设为目标值并下发, RTU(COM9) 跨传输回读断言;
物理背光亮灭无寄存器可读, 作 MANUAL 目视项随报告输出。

运行:
  pytest auto_testcase/test_acuview_Function_AcuRev1320_027_02_case1.py -v -s
前置: 活动桌面解锁; Add Connection 列表含测试表(行号见 config.test.yaml app.connection_row);
      1920x1080 / 125% / 最大化; 用例需连续跑(TCP 5min 空闲断连)。
"""
from __future__ import annotations

from pathlib import Path

import pytest

from comm.ctl_acuview.gui_driver import is_session_locked
from comm.ctl_acuview.testcase_engine import run_write_verify_case

# 用例在 projects/<项目>/tests/acuview_auto/，config 在项目根 → 上溯两级
PROJECT_ROOT = Path(__file__).resolve().parents[2]      # acuview_auto -> tests -> <项目>
TEST_CONFIG = str(PROJECT_ROOT / "config_acuview.yaml")

pytestmark = pytest.mark.skipif(
    is_session_locked(),
    reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面(截图与点击不可用)",
)


def test_case():
    """
    用例编号 : Function_AcuRev1320_027_02_case1
    用例标题 : HMI上设置背光延迟参数0成功，电表HMI无操作，电表背光不关闭。
    预置条件 : 1、Acuview2上位机  2、网线/RTU串口线  3、AcuRev1320 测试表
    测试步骤 : Setting>Metering>General 把 LCD Backlight Time 设为 0min 并 Update 下发
    预期结果 : 观察到 HMI 常亮背光(物理项, MANUAL); 自动判据: 0x1018 回读 == 0
    """
    report = run_write_verify_case(
        case_meta={
            "编号": "Function_AcuRev1320_027_02_case1",
            "标题": "HMI上设置背光延迟参数0成功，电表HMI无操作，电表背光不关闭。",
            "预置条件": "1、Acuview2上位机 2、网线/RTU串口线 3、AcuRev1320 测试表",
            "测试步骤": "General 页 LCD Backlight Time=0min 并 Update 下发",
            "预期结果": "HMI 常亮背光(MANUAL); 0x1018 回读==0",
        },
        register=4120, page="General", widget="Backlight_Value_Combo",
        target_value=0, physical_note="HMI 背光常亮(不关闭)", config_path=TEST_CONFIG,
    )
    assert report.passed, "背光时间未成功设为 0, 或回读不匹配 (详见 reports/auto_*.json)"
