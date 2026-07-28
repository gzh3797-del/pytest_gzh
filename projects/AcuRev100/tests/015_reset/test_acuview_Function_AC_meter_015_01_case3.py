r"""Function_AC meter_015_01_case3
用例标题: 上位机上通过RTU，恢复出厂设置并重启 (factory reset and reboot)
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: Reading->System Status 页点 Factory Reset 按钮 + 确认; 等掉线再重连恢复;
          经 Modbus 回读关键出厂默认值(SlaveID=1 / Frequency=1(60Hz) / Service=2 /
          CT Primary=1000 / Password=0)。
预期结果: 恢复出厂成功并重启, 基本参数回默认值, 重连后功能正常。

生成说明: run_button_action_case(is_reset=True)——点 Factory_Reset_Function_Button + 确认 +
等掉线/重连 + Modbus 校验出厂默认。
🔴🔴 红线操作(恢复出厂-最具破坏性, 重置全部配置): 用户本轮*未授权*自动执行, 仅生成挂门禁。
需工程师明确授权+在场后删除本 skip 运行(运行前需知悉将重置 SlaveID/波特率/密码/CT/接线/时间)。
⚠️ 依赖 COM11 + 桌面未锁 + 源保活。恢复出厂后密码=0/slaveID=1, 校验端按默认参数读。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_button_action_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

# 🔴🔴红线-恢复出厂(重置全部配置): 2026-07-15 用户在场明确授权放行("没有,赶紧干")。
#   跑完电表回出厂默认(SlaveID=1/Freq=60Hz/CT=1000/PW=0/USB Baud=19200/PhaseOrder=ABC)。


def test_015_01_case3():
    report = run_button_action_case(
        case_meta={
            '编号': 'Function_AC meter_015_01_case3',
            '标题': '上位机上通过RTU，恢复出厂设置并重启 (factory reset and reboot)',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'System Status 点 Factory Reset + 确认; 等掉线/重连; Modbus 校验出厂默认',
            '预期结果': '恢复出厂并重启成功, 基本参数回默认(SlaveID=1/Freq=1/Service=2/CT=1000/PW=0)',
        },
        page="System_Status", button_widget="Factory_Reset_Function_Button",
        is_reset=True,
        reset_verify=[
            ("出厂 SlaveID=1", 4111, 1),
            ("出厂 Frequency Selection=1(60Hz)", 4161, 1),
            ("出厂 CT Primary=1000", 4170, 1000),
            ("出厂 Password=0", 4096, 0),
        ],
        recover_timeout=150, config_path=TEST_CONFIG,
    )
    assert report.passed
