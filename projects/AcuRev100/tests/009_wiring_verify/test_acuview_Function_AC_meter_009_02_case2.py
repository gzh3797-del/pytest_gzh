r"""Function_AC meter_009_02_case2
用例标题: 上位机配2E3W1P接线， 系统功率数据是A、C相数据之和
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、Va、Vc=120V, A相50A(源5A)、C相150A(源15A), 压流夹角0°(Va/Vc角0/180°, 电流角0/180°)
3、系统平均电流100A, P_A+P_C=P_SYS
预期结果: Iavg=100A±0.2%; P_A≈6kW, P_C≈18kW, P_SYS≈24kW(±0.2%)

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_009_02_case2():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_009_02_case2',
        '标题': '上位机配2E3W1P接线， 系统功率数据是A、C相数据之和',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、Va、Vc=120V, A相50A(源5A)、C相150A(源15A), 压流夹角0°(Va/Vc角0/180°, 电流角0/180°)\n3、系统平均电流100A, P_A+P_C=P_SYS',
        '预期结果': 'Iavg=100A±0.2%; P_A≈6kW, P_C≈18kW, P_SYS≈24kW(±0.2%)',
    }, config_path=TEST_CONFIG)
    assert report.passed
