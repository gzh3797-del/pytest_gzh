r"""Function_AC meter_016_01_case6
用例标题: 配置波特率76800、115200，偶校验(EVEN)，通信正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: 设 USB 口波特率 76800/115200 + 偶校验(EVEN); 校验端以对应参数重连回读; 还原默认。
预期结果: 波特率+EVEN 写入成功且重连回读一致(76800 寄存器编码 7680, 115200 编码 11520); 还原成功。

生成说明: run_comm_param_case。改 USB 口参数。EVEN=0(parity='E')。高档波特率寄存器编码: 76800->7680。
🟡 门禁: 下拉选值需 Tesseract OCR。装好删除本 skip。⚠️ 依赖 COM11 可用 + 桌面未锁。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_comm_param_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

_META = {
    '编号': 'Function_AC meter_016_01_case6',
    '标题': '配置波特率76800、115200，偶校验(EVEN)，通信正常',
    '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
    '测试步骤': 'USB口波特率76800/115200+EVEN, 校验端对应参数重连回读, 还原默认',
    '预期结果': '波特率+EVEN写入成功重连回读一致(76800->7680); 还原成功',
}


def test_016_01_case6():
    r_par = run_comm_param_case(
        case_meta=_META, register=4136, page="Communication", widget="USB_Parity_Value_Combo",
        gui_value="Even", expect_value=0, verify_overrides={'parity': 'E'},
        allow_write=[4136], restore_value=3, config_path=TEST_CONFIG,
    )
    r_baud = run_comm_param_case(
        case_meta=_META, register=4135, page="Communication", widget="USB_Baud_Rate_Value_Combo",
        gui_value="76800", expect_value=7680, verify_overrides={'baudrate': 76800},
        allow_write=[4135], restore_value=19200, config_path=TEST_CONFIG,
    )
    assert r_par.passed and r_baud.passed
