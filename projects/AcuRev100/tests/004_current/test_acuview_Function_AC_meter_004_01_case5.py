r"""Function_AC meter_004_01_case5
用例标题: CT选型333mV(200A)，三路电流分别为(50A、100A，120A)，顺序ABC，相角设置为(0°、240°、120°)，电表设置电流分别为Positive、Negative，在上位机上分别配置为对应A相、B相、C相，测量Ia,Ib,Ic,Iavg满足精度0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、接线方式3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入三路电流分别为(1.25A、2.5A，3A)，在上位机上分别配置为对应A相、B相、C相
3、电表上设置为顺序ABC， 相角设置为(0°、240°、120°)
4、电表设置电流方向为Positive
5、Acuview2上检查交流电相电流是否均在精度范围内
6、电表设置电流方向为Negative
7、Acuview2上检查交流电相电流是否均在精度范围内
预期结果: 5、Acuview2上检查交流电相电流均在精度范围内
Ia: 49.900~50.100; Ib: 99.800~100.200; Ic: 119.760~120.240; Iavg: 89.820~90.180
7、Acuview2上检查交流电相电流均在精度范围内(幅值同上)
用户通道&输入通道&系统参数PF 为-1
用户通道&输入通道&系统参数有功功率P为负数

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
2026-07-14 用户确认: 两测点=同源(220V三相 + ia5/ib10/ic12A→显示50/100/120A), 点1方向
Positive, 点2方向 Negative(逐测点写 4168/4172/4176, spec 枚举 0:positive/1:negative);
点2 判据: 电流幅值同点1 + PF_A/B/C∈[-1.002,-0.998] + P_A/B/C 为负(V×I±10%)。
用例末强制还原方向=Positive(helpers 还原失败重试一次, 仍失败整例 FAIL 并置顶告警)。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_01_case5():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_01_case5',
        '标题': 'CT选型333mV(200A)，三路电流分别为(50A、100A，120A)，顺序ABC，相角设置为(0°、240°、120°)，电表设置电流分别为Positive、Negative，在上位机上分别配置为对应A相、B相、C相，测量Ia,Ib,Ic,Iavg满足精度0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、接线方式3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入三路电流分别为(1.25A、2.5A，3A)，在上位机上分别配置为对应A相、B相、C相\n3、电表上设置为顺序ABC， 相角设置为(0°、240°、120°)\n4、电表设置电流方向为Positive\n5、Acuview2上检查交流电相电流是否均在精度范围内\n6、电表设置电流方向为Negative\n7、Acuview2上检查交流电相电流是否均在精度范围内',
        '预期结果': '5、Acuview2上检查交流电相电流均在精度范围内\nIa：49.900~50.100\nIb: 99.800~100.200\nIc: 119.760~120.240\nIavg: 89.820~90.180\n7、Acuview2上检查交流电相电流均在精度范围内\nIa：49.900~50.100\nIb: 99.800~100.200\nIc: 119.760~120.240\nIavg: 89.820~90.180\n用户通道&输入通道&系统参数PF 为-1\n用户通道&输入通道&系统参数有功功率P为负数',
    }, config_path=TEST_CONFIG)
    assert report.passed
