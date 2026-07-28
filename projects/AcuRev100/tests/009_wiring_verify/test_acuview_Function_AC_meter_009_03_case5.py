r"""Function_AC meter_009_03_case5
用例标题: 上位机setting的Current&Wiring界面，修改CT Type与CT Primary（5-2000A连续可配）
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Current&Wiring 页遍历 CT Type 下拉(333mV/100mA/RCT)逐项保存回读;
          CT Primary 写 5(下限)/250/2000(上限)逐次保存回读; 还原 CT Primary=1000。
预期结果: CT Type 三项写入回读一致; CT Primary 5-2000 任意值写入成功回读一致; 还原=1000。

生成说明: run_multi_write_verify_case——CT Type 为 comboBox(选值依赖 OCR), CT Primary 为 lineEdit。
🟡 门禁两项(待解): ①下拉选值需 Tesseract OCR(明天装); ②CT Type 寄存器枚举值(0x1049=4169)地址表
仅给"0:100mA", 333mV/RCT 枚举索引待真机确认(项目 README 固件问题包在列, 与 CT_PRIMARY 连续/档位
差异同源)。真机确认枚举 + OCR 就绪后, 补全 _CT_TYPE_STEPS 并删除本 skip。
⚠️ 依赖 COM11(RS485) 可用 + 桌面未锁。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.testcase_engine import run_multi_write_verify_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = pytest.mark.skip(reason="🟡待Tesseract OCR + CT Type枚举真机确认(0x1049 仅知100mA=0)后启用")

# TODO(真机): CT Type 下拉文本 -> 寄存器枚举 int。地址表仅确认 100mA=0; 333mV/RCT 待确认。
_CT_TYPE_STEPS = [("100mA", 0)]  # ("333mV", ?), ("RCT", ?)
_CT_PRIMARY_STEPS = [5, 250, 2000]


def test_009_03_case5():
    _META = {
        '编号': 'Function_AC meter_009_03_case5',
        '标题': '上位机setting的Current&Wiring界面，修改CT Type与CT Primary（5-2000A连续可配）',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表(铅封解锁)',
        '测试步骤': 'CT Type 遍历 333mV/100mA/RCT 回读; CT Primary 写 5/250/2000 回读; 还原 1000',
        '预期结果': 'CT Type/CT Primary 写入回读一致; 还原=1000',
    }
    # CT Type 遍历(comboBox @0x1049=4169)
    r1 = run_multi_write_verify_case(
        case_meta=_META, register=4169, page="Current_Wire", widget="Channel_A_CT_Type_Combo",
        steps=_CT_TYPE_STEPS, restore_value=0, config_path=TEST_CONFIG,
    )
    # CT Primary 多点(lineEdit @0x104A=4170)
    r2 = run_multi_write_verify_case(
        case_meta=_META, register=4170, page="Current_Wire", widget="Channel_A_CT_Primary_LineEdit",
        steps=_CT_PRIMARY_STEPS, restore_value=1000, config_path=TEST_CONFIG,
    )
    assert r1.passed and r2.passed
