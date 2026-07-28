r"""Function_AC meter_017_01_case1
用例标题: 首次连接上位机进入general，修改参数(密码权限管理)
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 进 General 改参数点 Update -> 输错密码(提示错误) -> 输对密码(0000, 修改成功) ->
          切 reading 再进 general 改参数点 Update(本次无需密码)。
预期结果: 错误密码提示错误; 正确密码修改成功; 同一连接内再次修改无需密码(重连后仍需)。

实现: run_password_gate_case(用例类型7)。密码门禁=每连接首个 Setting Update 需密码(弹 "Please Enter
  Password" 预填0000), 输对一次后本连接免密, 关连接重开重置。可逆锚点=Energy Pulse Constant
  (0x1066=4198, Modbus 可写, GUI值×1000=寄存器), 以 Modbus 回读判据: 输错→回读=原值(未写);
  输对→回读=first_expect; 免密再改→回读=second_expect。arm_gate=True 先重连武装门禁。

已验证(2026-07-15): 核心三态在已武装连接上实跑全 PASS(输错拒绝/输对生效/免密再改/还原)。
  重连(arm_gate)按用户确认手法实现(标签×→Yes→Yes→Add Connection→Connect), 首次真机跑若坐标需
  微调即调 _RECONNECT_XY。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_password_gate_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_017_01_case1():
    report = run_password_gate_case(
        case_meta={
            '编号': 'Function_AC meter_017_01_case1',
            '标题': '首次连接上位机进入general，修改参数(密码权限管理)',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
            '测试步骤': 'General改参Update: 输错密码提示错误; 输对0000生效; 同连接再改免密',
            '预期结果': '错误密码提示错误; 正确密码修改成功; 同一连接内再次修改无需密码',
        },
        page="General", widget="Energy_Pulse_Constant_Value_Edit", anchor_register=4198,
        first_gui="2", first_expect=2000, second_gui="3", second_expect=3000,
        restore_reg=1000, wrong_password="9999", correct_password="0000",
        arm_gate=True, config_path=TEST_CONFIG,
    )
    assert report.passed
