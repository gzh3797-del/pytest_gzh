r"""Function_AC meter_007_01_case5
用例标题: CT选型100mA(200A)，3E4WY接线，输入ABC电压均为270V，(A&B&C相)电流为(200A、200A、200A)，电压角度分别为(0、240、120)，电流角度分别为(60、300、180)，AC meter上检查能量累计Ep、Eq、Es精度
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、源输入A/B相20A、C相15A(20A裁决降流)，电压均为270V
3、电压角(0、240、120)，电流角(60、300、180)容性
4、(自动化)记录能量基线代替清零
6、累计时间T=10分钟后读取能量增量Δ
7、检查Δ精度是否满足±0.5%要求
预期结果: Ep+ A/B 4.478~4.523, C 3.358~3.393, 系12.348~12.402 kWh; Eq- A/B 7.755~7.833, C 5.816~5.875, 系21.387~21.482 kvarh(EXPORT); Es A/B 8.982~9.018, C 6.736~6.764, 系24.696~24.804 kVAh

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 能量增量断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_energy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_007_01_case5():
    report = run_energy_case(case_meta={
        '编号': 'Function_AC meter_007_01_case5',
        '标题': 'CT选型100mA(200A)，3E4WY接线，输入ABC电压均为270V，(A&B&C相)电流为(200A、200A、200A)，电压角度分别为(0、240、120)，电流角度分别为(60、300、180)，AC meter上检查能量累计Ep、Eq、Es精度',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、源输入A/B相20A、C相15A(20A裁决降流)，电压均为270V\n3、电压角(0、240、120)，电流角(60、300、180)容性\n4、(自动化)记录能量基线代替清零\n6、累计时间T=10分钟后读取能量增量Δ\n7、检查Δ精度是否满足±0.5%要求',
        '预期结果': 'Ep+ A/B 4.478~4.523, C 3.358~3.393, 系12.348~12.402 kWh; Eq- A/B 7.755~7.833, C 5.816~5.875, 系21.387~21.482 kvarh(EXPORT); Es A/B 8.982~9.018, C 6.736~6.764, 系24.696~24.804 kVAh',
    }, config_path=TEST_CONFIG)
    assert report.passed
