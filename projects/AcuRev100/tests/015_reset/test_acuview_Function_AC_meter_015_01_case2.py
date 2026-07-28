r"""Function_AC meter_015_01_case2
用例标题: 上位机上通过RTU，"Reading"->"system status"->"Reboot Meter"可reboot电表，重启后功能正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: Reading->System Status 页点 Reboot Meter 按钮 + 确认; 等电表掉线再重连恢复;
          确认重连后配置寄存器可读、电气测量正常(测量部分依赖 ADC)。
预期结果: 电表重启成功, 重连后功能正常(配置在线可读; 电气测量满足精度=换板后有效)。

生成说明: run_button_action_case(is_reset=True)——点 Meter_Reboot_Function_Button + 确认弹窗 +
等掉线(_wait_offline)+等重连恢复(_wait_online)。
🔴 红线操作(重启自供电表): 已获用户授权, 但无人值守时源停0卡死无法现场恢复 → 必须工程师在场
手动放行(删除本 skip)后运行。"电气测量正常"部分当前板 ADC 损坏为已知 FAIL(MANUAL)。
⚠️ 依赖 COM11(RS485) 可用 + 桌面未锁 + CL3021 源保活(重启后自供电靠 Va 复活)。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_button_action_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

# 🔴红线-重启自供电表: 2026-07-15 用户在场明确确认放行("可以,直接干")。重启后自供电靠 Va 复活,
#   CL3021 源在场保活。"电气测量正常"部分当前板 ADC 损坏为已知 FAIL(MANUAL)。


def test_015_01_case2():
    report = run_button_action_case(
        case_meta={
            '编号': 'Function_AC meter_015_01_case2',
            '标题': '上位机 Reading->system status->Reboot Meter 重启电表，重启后功能正常',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、铅封解锁',
            '测试步骤': 'System Status 点 Reboot Meter + 确认; 等掉线再重连恢复; 确认功能正常',
            '预期结果': '重启成功, 重连后配置在线可读(测量满足精度=换板后有效)',
        },
        page="System_Status", button_widget="Meter_Reboot_Function_Button",
        is_reset=True,   # elevate_first 默认 False: 密码仅在真弹框时由 _confirm_dialogs 处理(每连接一次性)
        reset_verify=[("重启后 SERVICE_CONFIGURATION 可读(配置保持)", 4162, 2)],
        recover_timeout=120, config_path=TEST_CONFIG,
    )
    assert report.passed
