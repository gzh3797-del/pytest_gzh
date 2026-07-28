r"""Function_AC meter_007_01_case2
用例标题: CT选型333mV(200A)，1E2W接线，输入A相电压均为277V，A相电流为100A，电压角度0°，电流角度为120°，AC meter上检查能量累计Ep、Eq、Es精度
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、源输入A相电压277V、电流2.5A
3、电压角0°, 电流角120°(P<0发电,容性)
4、(自动化)记录能量基线代替清零
6、累计时间T=10分钟后读取能量增量Δ
7、检查Δ精度是否满足±0.5%要求
预期结果: Ep-(发电) A/系 2.297~2.320 kWh(EXPORT); Eq-(容性) A/系 3.978~4.018 kvarh(EXPORT); Es A/系 4.594~4.640 kVAh

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 能量增量断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_energy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_007_01_case2():
    report = run_energy_case(case_meta={
        '编号': 'Function_AC meter_007_01_case2',
        '标题': 'CT选型333mV(200A)，1E2W接线，输入A相电压均为277V，A相电流为100A，电压角度0°，电流角度为120°，AC meter上检查能量累计Ep、Eq、Es精度',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、源输入A相电压277V、电流2.5A\n3、电压角0°, 电流角120°(P<0发电,容性)\n4、(自动化)记录能量基线代替清零\n6、累计时间T=10分钟后读取能量增量Δ\n7、检查Δ精度是否满足±0.5%要求',
        '预期结果': 'Ep-(发电) A/系 2.297~2.320 kWh(EXPORT); Eq-(容性) A/系 3.978~4.018 kvarh(EXPORT); Es A/系 4.594~4.640 kVAh',
    }, config_path=TEST_CONFIG)
    assert report.passed
