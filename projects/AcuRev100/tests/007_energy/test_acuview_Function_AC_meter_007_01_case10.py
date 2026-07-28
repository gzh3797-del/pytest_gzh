r"""Function_AC meter_007_01_case10
用例标题: CT选型100mA(200A)，3E4WY接线，三相电压均为220V，三相电流不平衡（A相50A、B相100A、C相150A），三相电压角度(0、240、120)，三相电流角度(330、210、90)（各相φ均为30°，感性），AC meter上检查各相及系统能量累计Ep、Eq、Es精度
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、源输入三相电压220V；A相5A、B相10A、C相15A(不平衡)
3、电压角(0、240、120)，电流角(330、210、90)即各相φ=30°感性
4、(自动化)记录能量基线代替清零
6、累计时间T=10分钟后读取能量增量Δ
7、检查Δ精度是否满足±0.5%要求
预期结果: Ep+ A1.580~1.596/B3.160~3.191/C4.739~4.787/系9.507~9.545 kWh; Eq+ A0.912~0.921/B1.824~1.842/C2.736~2.764/系5.489~5.511 kvarh; Es A1.824~1.842/B3.648~3.685/C5.489~5.511/系10.978~11.022 kVAh

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 能量增量断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_energy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_007_01_case10():
    report = run_energy_case(case_meta={
        '编号': 'Function_AC meter_007_01_case10',
        '标题': 'CT选型100mA(200A)，3E4WY接线，三相电压均为220V，三相电流不平衡（A相50A、B相100A、C相150A），三相电压角度(0、240、120)，三相电流角度(330、210、90)（各相φ均为30°，感性），AC meter上检查各相及系统能量累计Ep、Eq、Es精度',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、源输入三相电压220V；A相5A、B相10A、C相15A(不平衡)\n3、电压角(0、240、120)，电流角(330、210、90)即各相φ=30°感性\n4、(自动化)记录能量基线代替清零\n6、累计时间T=10分钟后读取能量增量Δ\n7、检查Δ精度是否满足±0.5%要求',
        '预期结果': 'Ep+ A1.580~1.596/B3.160~3.191/C4.739~4.787/系9.507~9.545 kWh; Eq+ A0.912~0.921/B1.824~1.842/C2.736~2.764/系5.489~5.511 kvarh; Es A1.824~1.842/B3.648~3.685/C5.489~5.511/系10.978~11.022 kVAh',
    }, config_path=TEST_CONFIG)
    assert report.passed
