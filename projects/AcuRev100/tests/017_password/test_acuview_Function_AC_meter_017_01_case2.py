r"""Function_AC meter_017_01_case2
用例标题: 首次连接上位机进入Current&Wiring，修改参数(密码权限管理)
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 进 Current&Wiring 改参数点 Update -> 输错密码(提示错误) -> 输对密码(修改成功) ->
          切 reading 再进 general 改参数点 Update(本次无需密码)。
预期结果: 错误密码提示错误; 正确密码修改成功; 同一连接内再次修改无需密码。

实现: run_password_gate_case(同 017_01_case1, 页面换 Current&Wiring)。可逆锚点=Channel A CT Primary
  (0x104A=4170, R/W, 无缩放, 合法 5-2000, lineEdit 键入不依赖 OCR), Modbus 回读判据。
  arm_gate 重连武装门禁; restore_reg=None → 还原到用例开始时的原值。密码门禁机制见 017_01_case1。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_password_gate_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_017_01_case2():
    report = run_password_gate_case(
        case_meta={
            '编号': 'Function_AC meter_017_01_case2',
            '标题': '首次连接上位机进入Current&Wiring，修改参数(密码权限管理)',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
            '测试步骤': 'Current&Wiring改CT Primary Update: 输错提示错误; 输对生效; 同连接再改免密',
            '预期结果': '错误密码提示错误; 正确密码修改成功; 同一连接内再次修改无需密码',
        },
        page="Current_Wire", widget="Channel_A_CT_Primary_LineEdit", anchor_register=4170,
        first_gui="100", first_expect=100, second_gui="200", second_expect=200,
        restore_reg=None, wrong_password="9999", correct_password="0000",
        arm_gate=True, config_path=TEST_CONFIG,
    )
    assert report.passed
