r"""Function_AC meter_009_03_case3
用例标题: 3E4WY接线配置，上位机 realtime数据中“显示”符合接线方式中的显示规划
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、3E4WY接线, V=120V×3, I_in1~3原文200A×3 → 源20/20/15A(C相降流)
6、real-time 可测寄存器断言; user1/显示规划转MANUAL
预期结果: V≈120V×3, Ia/Ib≈200A, Ic≈150A, P_A/P_B≈24kW, P_C≈18kW, P_SYS≈66kW(±0.2%)

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_009_03_case3():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_009_03_case3',
        '标题': '3E4WY接线配置，上位机 realtime数据中“显示”符合接线方式中的显示规划',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、3E4WY接线, V=120V×3, I_in1~3原文200A×3 → 源20/20/15A(C相降流)\n6、real-time 可测寄存器断言; user1/显示规划转MANUAL',
        '预期结果': 'V≈120V×3, Ia/Ib≈200A, Ic≈150A, P_A/P_B≈24kW, P_C≈18kW, P_SYS≈66kW(±0.2%)',
    }, config_path=TEST_CONFIG)
    assert report.passed
