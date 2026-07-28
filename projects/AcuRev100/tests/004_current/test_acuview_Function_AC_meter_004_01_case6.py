r"""Function_AC meter_004_01_case6
用例标题: CT选型333mV(400A)，3E4WY，三路电流分别为(50A、200A，400A)，在上位机上分别配置为对应A相、B相、C相，测量Ia,Ib,Ic,Iavg满足精度0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、接线方式3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入三路电流分别为(0.625A、2.5A，5A)，在上位机上分别配置为对应A相、B相、C相
3、Acuview2上检查交流电相电流是否均在精度范围内
预期结果: 3、Acuview2上检查交流电相电流均在精度范围内
Ia：	49.900	~	50.100
Ib:	199.600	~	200.400
Ic:	399.200	~	400.800
Iavg:	216.233	~	217.100

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_01_case6():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_01_case6',
        '标题': 'CT选型333mV(400A)，3E4WY，三路电流分别为(50A、200A，400A)，在上位机上分别配置为对应A相、B相、C相，测量Ia,Ib,Ic,Iavg满足精度0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、接线方式3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入三路电流分别为(0.625A、2.5A，5A)，在上位机上分别配置为对应A相、B相、C相\n3、Acuview2上检查交流电相电流是否均在精度范围内',
        '预期结果': '3、Acuview2上检查交流电相电流均在精度范围内\nIa：\t49.900\t~\t50.100\nIb:\t199.600\t~\t200.400\nIc:\t399.200\t~\t400.800\nIavg:\t216.233\t~\t217.100',
    }, config_path=TEST_CONFIG)
    assert report.passed
