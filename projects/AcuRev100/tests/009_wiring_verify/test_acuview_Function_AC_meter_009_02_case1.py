r"""Function_AC meter_009_02_case1
用例标题: 上位机配2E3W1P接线，压流夹角30°
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、2E3W1P接线, Va、Vc=120V, A、C相电流50A(源5A), 压流夹角30°
3、检查A/C/系统功率、系统PF=0.866
4、系统平均电流50A
预期结果: P_A=P_C≈5196W, P_SYS≈10392W(±0.2%), PF_SYS=0.866±0.002, Iavg=50A±0.2%

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_009_02_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_009_02_case1',
        '标题': '上位机配2E3W1P接线，压流夹角30°',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、2E3W1P接线, Va、Vc=120V, A、C相电流50A(源5A), 压流夹角30°\n3、检查A/C/系统功率、系统PF=0.866\n4、系统平均电流50A',
        '预期结果': 'P_A=P_C≈5196W, P_SYS≈10392W(±0.2%), PF_SYS=0.866±0.002, Iavg=50A±0.2%',
    }, config_path=TEST_CONFIG)
    assert report.passed
