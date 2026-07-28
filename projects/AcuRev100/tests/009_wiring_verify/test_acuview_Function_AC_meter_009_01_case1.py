r"""Function_AC meter_009_01_case1
用例标题: 1E2W 接线方式， 上位机 realtime数据中“显示”符合接线方式中的显示规划
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、1E2W输电,上位机配1E2W接线，Va为120V, I_in1为100A(源10A经台体CT)，压流夹角0°
3、real-time 可测寄存器区间断言; '-'显示项转MANUAL
预期结果: Va≈120V, Ia≈100A, P_A=P_SYS≈12kW(±0.2%); '-'项与Energy页显示规划 MANUAL 目视

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_009_01_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_009_01_case1',
        '标题': '1E2W 接线方式， 上位机 realtime数据中“显示”符合接线方式中的显示规划',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、1E2W输电,上位机配1E2W接线，Va为120V, I_in1为100A(源10A经台体CT)，压流夹角0°\n3、real-time 可测寄存器区间断言; \'-\'显示项转MANUAL',
        '预期结果': 'Va≈120V, Ia≈100A, P_A=P_SYS≈12kW(±0.2%); \'-\'项与Energy页显示规划 MANUAL 目视',
    }, config_path=TEST_CONFIG)
    assert report.passed
