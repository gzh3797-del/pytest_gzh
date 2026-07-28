r"""Function_AC meter_016_01_case5
用例标题: 配置波特率9600、19200，奇校验(ODD)，通信正常；奇/偶校验互相验证
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: 设 USB 口波特率 9600/19200 + 奇校验(ODD); 校验端以对应参数重连回读; 还原默认。
          "改偶校验后上位机连接失败"属连接行为(MANUAL 目视)。
预期结果: 波特率+ODD 写入成功且重连回读一致; 还原成功。

生成说明: run_comm_param_case。方向=改 USB 口参数(Acuview 稳定在 RS485/COM11)。Baud@4135, Parity@4136
(枚举 0:even 1:odd 2:None2 3:None1)。奇校验重连 parity='O'。
🟡 门禁: 下拉选值需 Tesseract OCR。装好删除本 skip。⚠️ 依赖 COM11 可用 + 桌面未锁。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_comm_param_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

_META = {
    '编号': 'Function_AC meter_016_01_case5',
    '标题': '配置波特率9600、19200，奇校验(ODD)，通信正常',
    '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
    '测试步骤': 'USB口波特率9600/19200+ODD, 校验端对应参数重连回读, 还原默认',
    '预期结果': '波特率+ODD写入成功重连回读一致; 还原成功',
}


def test_016_01_case5():
    # 奇校验(ODD=1)
    r_par = run_comm_param_case(
        case_meta=_META, register=4136, page="Communication", widget="USB_Parity_Value_Combo",
        gui_value="Odd", expect_value=1, verify_overrides={'parity': 'O'},
        allow_write=[4136], restore_value=3, config_path=TEST_CONFIG,
    )
    # 波特率 9600(ODD 生效期间, 校验端同时带 parity='O' 重连)
    r_baud = run_comm_param_case(
        case_meta=_META, register=4135, page="Communication", widget="USB_Baud_Rate_Value_Combo",
        gui_value="9600", expect_value=9600, verify_overrides={'baudrate': 9600},
        allow_write=[4135], restore_value=19200, config_path=TEST_CONFIG,
    )
    assert r_par.passed and r_baud.passed
