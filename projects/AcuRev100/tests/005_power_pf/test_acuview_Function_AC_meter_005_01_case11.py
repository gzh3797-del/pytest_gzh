r"""Function_AC meter_005_01_case11
用例标题: CT选型RCT(500A)，3E4WY接线，输入ABC电压均为100V，(A&B&C相)电流为(200A、300A、400A)，电压角度分别为(0，240，120)，电流角度分别为(0，240，120)，AC meter上检查测量PF、功率
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入ABC电压均为100V，(A&B&C相)电流为(2A、3A、4A)
3、电压角度分别为(0，240，120)，电流角度分别为(0，240，120)
4、AC meter上检查测量PF、功率精度是否满足要求
预期结果: 4、AC meter上检查测量PF、功率精度满足要求
有功功率 P 取值（±0.2%，工程设计目标；认证硬指标为0.5%）：
A相：19960.0 ~ 20040.0 W
B相：29940.0 ~ 30060.0 W
C相：39920.0 ~ 40080.0 W
系统：89820.0 ~ 90180.0 W
无功功率 Q 取值（±0.5%，以额定S为基准，SRS §3.5）：
A相：-100.0 ~ 100.0 VAR
B相：-150.0 ~ 150.0 VAR
C相：-200.0 ~ 200.0 VAR
系统：-450.0 ~ 450.0 VAR
视在功率 S 取值（±0.5%，SRS §3.5）：
A相：19900.0 ~ 20100.0 VA
B相：29850.0 ~ 30150.0 VA
C相：39800.0 ~ 40200.0 VA
系统：89550.0 ~ 90450.0 VA
功率因数 PF 取值：（Class 0.2，±0.002）
A相：PF 0.998 ~ 1.002
B相：PF 0.998 ~ 1.002
C相：PF 0.998 ~ 1.002
系统：PF 0.998 ~ 1.002

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_005_01_case11():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_005_01_case11',
        '标题': 'CT选型RCT(500A)，3E4WY接线，输入ABC电压均为100V，(A&B&C相)电流为(200A、300A、400A)，电压角度分别为(0，240，120)，电流角度分别为(0，240，120)，AC meter上检查测量PF、功率',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入ABC电压均为100V，(A&B&C相)电流为(2A、3A、4A)\n3、电压角度分别为(0，240，120)，电流角度分别为(0，240，120)\n4、AC meter上检查测量PF、功率精度是否满足要求',
        '预期结果': '4、AC meter上检查测量PF、功率精度满足要求\n有功功率 P 取值（±0.2%，工程设计目标；认证硬指标为0.5%）：\nA相：19960.0 ~ 20040.0 W\nB相：29940.0 ~ 30060.0 W\nC相：39920.0 ~ 40080.0 W\n系统：89820.0 ~ 90180.0 W\n无功功率 Q 取值（±0.5%，以额定S为基准，SRS §3.5）：\nA相：-100.0 ~ 100.0 VAR\nB相：-150.0 ~ 150.0 VAR\nC相：-200.0 ~ 200.0 VAR\n系统：-450.0 ~ 450.0 VAR\n视在功率 S 取值（±0.5%，SRS §3.5）：\nA相：19900.0 ~ 20100.0 VA\nB相：29850.0 ~ 30150.0 VA\nC相：39800.0 ~ 40200.0 VA\n系统：89550.0 ~ 90450.0 VA\n功率因数 PF 取值：（Class 0.2，±0.002）\nA相：PF 0.998 ~ 1.002\nB相：PF 0.998 ~ 1.002\nC相：PF 0.998 ~ 1.002\n系统：PF 0.998 ~ 1.002',
    }, config_path=TEST_CONFIG)
    assert report.passed
