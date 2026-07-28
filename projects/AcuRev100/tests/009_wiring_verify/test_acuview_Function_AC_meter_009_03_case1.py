r"""Function_AC meter_009_03_case1
用例标题: 上位机配3E4WY接线，A\B\C三相电流分别配Va、Vb、Vc，压流角0°，对应功率正确
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、3E4WY接线, Va/Vb/Vc=120V, I_in1~3=50/100/150A(源5/10/15A), 压流角0°
4、real-time 数据区间断言
预期结果: V≈120V×3, I≈50/100/150A, P≈6/12/18kW, P_SYS≈36kW(±0.2%)

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_009_03_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_009_03_case1',
        '标题': '上位机配3E4WY接线，A\\B\\C三相电流分别配Va、Vb、Vc，压流角0°，对应功率正确',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、3E4WY接线, Va/Vb/Vc=120V, I_in1~3=50/100/150A(源5/10/15A), 压流角0°\n4、real-time 数据区间断言',
        '预期结果': 'V≈120V×3, I≈50/100/150A, P≈6/12/18kW, P_SYS≈36kW(±0.2%)',
    }, config_path=TEST_CONFIG)
    assert report.passed
