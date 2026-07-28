r"""Function_AC meter_016_01_case4
用例标题: 通过Rtu/USB,配置modbus RTU ON成功，波特率9600、None1，通信正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Acuview(RS485/COM11)在 Communication 页设 USB 口波特率=9600(None1 为默认校验);
          校验端(USB/COM6)以 9600 重连回读 USB Baud(0x1027=4135)=9600; 还原=19200。
          "以新波特率经 RTU/USB 连接成功"属上位机连接行为(MANUAL 目视)。
预期结果: USB 波特率写 9600 成功且以 9600 重连回读一致; 还原 19200。

生成说明: run_comm_param_case。**清路径**: Acuview 固定 RS485(COM11)下发, 改的是 *USB口* 参数
(4135/4136)——不影响 Acuview 自身 COM11 链路; 校验端 COM6 用新波特率重连回读。改 RS485 口参数会断
Acuview 链路, 故 016 波特率/校验统一走"改USB口"方向。波特率高档寄存器编码: 76800->7680, 115200->11520。
🟡 门禁: 下拉选值需 Tesseract OCR(明天装)。装好删除本 skip。
⚠️ 依赖 COM11 可用 + 桌面未锁。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_comm_param_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_016_01_case4():
    report = run_comm_param_case(
        case_meta={
            '编号': 'Function_AC meter_016_01_case4',
            '标题': '通过Rtu/USB,配置modbus波特率9600、None1，通信正常',
            '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
            '测试步骤': 'USB口波特率设 9600(None1默认); COM6 以 9600 重连回读 4135; 还原 19200',
            '预期结果': 'USB波特率写9600成功重连回读一致; 还原19200',
        },
        register=4135, page="Communication", widget="USB_Baud_Rate_Value_Combo",
        gui_value="9600", expect_value=9600, verify_overrides={'baudrate': 9600},
        allow_write=[4135], restore_value=19200, config_path=TEST_CONFIG,
    )
    assert report.passed
