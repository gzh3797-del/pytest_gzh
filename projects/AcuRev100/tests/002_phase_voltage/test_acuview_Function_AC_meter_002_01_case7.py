r"""Function_AC meter_002_01_case7
用例标题: 电压为0，只加电流, 电流测量值正常，但不计算夹角、功率、能量；电压为10V，电流测量值正常，且计算夹角、功率、能量
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的额定电压为0V。
3、电表相序设置为ABC，φA、φB、φC分别为(0°、240°、120°)
4、input1-3 设置为3A
5、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值是否为0V。
6、ABC三相, 输入通道input1-3电流均为3A
7、交流电源输入相电压Van、Vbn、Vcn的额定电压为10V。
8、电表相序设置为ABC，φA、φB、φC分别为(0°、240°、120°)
9、input1-3 设置为3A
10、ABC三相, 输入通道input1-3电流均为3A
预期结果: 5、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值为0V
6、ABC三相, 输入通道input1-3、均为3A，同时功率、夹角、能量不计算
10、ABC三相, 输入通道input1-3、电流均为3A，同时计算功率、夹角、能量

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
2026-07-14 用户裁决(自供电冲突解法): A相常驻最低供电 100V(config
source.supply_guard 可配, 代码 _guard_point 兜底), 0V/10V 只加 Ub/Uc; 判据仅
B/C 相: 电流 Ib/Ic=3A±10%(源0.3A×台体CT系数10), 0V点 Pb=Pc=0, 10V点 Pb=Pc=30W±10%。
夹角/能量不计入自动断言(0V下夹角寄存器行为未知; 能量由005系列覆盖) → MANUAL。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_002_01_case7():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_002_01_case7',
        '标题': '电压为0，只加电流, 电流测量值正常，但不计算夹角、功率、能量；电压为10V，电流测量值正常，且计算夹角、功率、能量',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的额定电压为0V。\n3、电表相序设置为ABC，φA、φB、φC分别为(0°、240°、120°)\n4、input1-3 设置为3A\n5、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值是否为0V。\n6、ABC三相, 输入通道input1-3电流均为3A\n7、交流电源输入相电压Van、Vbn、Vcn的额定电压为10V。\n8、电表相序设置为ABC，φA、φB、φC分别为(0°、240°、120°)\n9、input1-3 设置为3A\n10、ABC三相, 输入通道input1-3电流均为3A',
        '预期结果': '5、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值为0V\n6、ABC三相, 输入通道input1-3、均为3A，同时功率、夹角、能量不计算\n10、ABC三相, 输入通道input1-3、电流均为3A，同时计算功率、夹角、能量',
    }, config_path=TEST_CONFIG)
    assert report.passed
