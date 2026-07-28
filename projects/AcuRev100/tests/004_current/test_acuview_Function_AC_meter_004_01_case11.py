r"""Function_AC meter_004_01_case11
用例标题: CT选型100mA(400A)，三路电流分别为Imin（4A，4A，4A）,在上位机上分别配置为对应A相、B相、C相，精度保证的电流范围0.2%,输入三路电流分别为（0.4A，0A，0.4A）
备注：0.4A为Ist
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、接线方式3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入三路电流分别为（0.2，0.2，0.2）,在上位机上分别配置为对应A相、B相、C相
3、精度保证的电流范围是否符合精度的要求
4、输入三路电流分别为（0.02，0，0.02）,在上位机上分别配置为对应A相、B相、C相
预期结果: 3、Acuview2上检查交流电相电流均在精度范围内
Ia：	3.992	~	4.008
Ib:	3.992	~	4.008
Ic:	3.992	~	4.008
Iavg:	3.992	~	4.008
4、Acuview2上检查交流电相电流分别为：（0.4A，0A，0.4A）(无精度要求)

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_01_case11():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_01_case11',
        '标题': 'CT选型100mA(400A)，三路电流分别为Imin（4A，4A，4A）,在上位机上分别配置为对应A相、B相、C相，精度保证的电流范围0.2%,输入三路电流分别为（0.4A，0A，0.4A）\n备注：0.4A为Ist',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、接线方式3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入三路电流分别为（0.2，0.2，0.2）,在上位机上分别配置为对应A相、B相、C相\n3、精度保证的电流范围是否符合精度的要求\n4、输入三路电流分别为（0.02，0，0.02）,在上位机上分别配置为对应A相、B相、C相',
        '预期结果': '3、Acuview2上检查交流电相电流均在精度范围内\nIa：\t3.992\t~\t4.008\nIb:\t3.992\t~\t4.008\nIc:\t3.992\t~\t4.008\nIavg:\t3.992\t~\t4.008\n4、Acuview2上检查交流电相电流分别为：（0.4A，0A，0.4A）(无精度要求)',
    }, config_path=TEST_CONFIG)
    assert report.passed
