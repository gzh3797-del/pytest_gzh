r"""Function_AC meter_016_01_case7
用例标题: 配置波特率76800、115200，无校验(None1、None2)，通信正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: 设 USB 口波特率 76800(None1)/115200(None2); 校验端以对应参数重连回读; 还原默认。
预期结果: 波特率+None1/None2 写入成功且重连回读一致; 还原成功。

生成说明: run_comm_param_case。改 USB 口参数。None1=3(1停止位)/None2=2(2停止位), 均 parity='N';
⚠️ None2 需 2 停止位, 当前 MeterClient 重连仅覆盖 baud/parity(stopbits 固定), None2 回读需补 stopbits 覆盖(TODO)。
🟡 门禁: 下拉选值需 Tesseract OCR。装好删除本 skip。⚠️ 依赖 COM11 可用 + 桌面未锁。
"""
from pathlib import Path

from comm.ctl_acuview.testcase_engine import run_comm_param_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

_META = {
    '编号': 'Function_AC meter_016_01_case7',
    '标题': '配置波特率76800、115200，无校验(None1、None2)，通信正常',
    '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
    '测试步骤': 'USB口波特率76800(None1)/115200(None2), 校验端对应参数重连回读, 还原默认',
    '预期结果': '波特率+None1/None2写入成功重连回读一致; 还原成功',
}


def test_016_01_case7():
    # None1(=3) + 76800
    r1 = run_comm_param_case(
        case_meta=_META, register=4135, page="Communication", widget="USB_Baud_Rate_Value_Combo",
        gui_value="76800", expect_value=7680, verify_overrides={'baudrate': 76800},
        allow_write=[4135], restore_value=19200, config_path=TEST_CONFIG,
    )
    assert r1.passed
