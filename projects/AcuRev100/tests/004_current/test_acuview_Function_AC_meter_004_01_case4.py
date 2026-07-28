r"""Function_AC meter_004_01_case4
用例标题: CT选型333mV(200A)，两路电流分别为(200A、240A),在上位机均配置为对应B相和C相，测量Ia,Ib,Ic,Iavg满足精度0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、接线方式3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入两路电流分别为(5A、6A),在上位机均配置为对应B相和C相。
3、Acuview2上检查交流电相电流是否均在精度范围内
预期结果: 3、Acuview2上检查交流电相电流均在精度范围内
Ia: 0~0; Ib: 199.600~200.400; Ic: 239.520~240.480; Iavg: 146.373~146.960

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
2026-07-14 源点补齐(原xlsx解析丢源点): 手工5A/6A按 mV→mA 换算×4 → ib=20A/ic=24A
(显示200A/240A, 低于 via_ct 25A 硬限幅), ia=0, 电压 220V 三相保活。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_01_case4():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_01_case4',
        '标题': 'CT选型333mV(200A)，两路电流分别为(200A、240A),在上位机均配置为对应B相和C相，测量Ia,Ib,Ic,Iavg满足精度0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、接线方式3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入两路电流分别为(5A、6A),在上位机均配置为对应B相和C相。\n3、Acuview2上检查交流电相电流是否均在精度范围内',
        '预期结果': '3、Acuview2上检查交流电相电流均在精度范围内\nIa：0~0\nIb: 199.600~200.400\nIc: 239.520~240.480\nIavg: 146.373~146.960',
    }, config_path=TEST_CONFIG)
    assert report.passed
