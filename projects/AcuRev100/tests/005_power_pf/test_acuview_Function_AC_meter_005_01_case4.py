r"""Function_AC meter_005_01_case4
用例标题: CT选型333mV(200A)，1E2W接线下，输入A相电压均为277V，A相电流为100A，电压角度0°，电流角度为120°，AC meter上检查测量PF、功率
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、1E2W
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、输入A相电压均为277V，A相电流为2.5A
3、电压角度0°，电流角度为120°
4、AC meter上检查测量PF、功率精度是否满足要求
预期结果: 4、AC meter上检查测量PF、功率精度满足要求
有功功率 P 取值（±0.2%，工程设计目标；认证硬指标为0.5%）：
A相：-13877.7 ~ -13822.2 W
系统：-13877.7 ~ -13822.2 W
无功功率 Q 取值（±0.5%，以额定S为基准，SRS §3.5）：
A相：-24127.4 ~ -23850.4 VAR
系统：-24127.4 ~ -23850.4 VAR
视在功率 S 取值（±0.5%，SRS §3.5）：
A相：27561.5 ~ 27838.5 VA
系统：27561.5 ~ 27838.5 VA
功率因数 PF 取值：（Class 0.2，±0.002）
A相：PF -0.502 ~ -0.498
系统：PF -0.502 ~ -0.498

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_005_01_case4():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_005_01_case4',
        '标题': 'CT选型333mV(200A)，1E2W接线下，输入A相电压均为277V，A相电流为100A，电压角度0°，电流角度为120°，AC meter上检查测量PF、功率',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、1E2W',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、输入A相电压均为277V，A相电流为2.5A\n3、电压角度0°，电流角度为120°\n4、AC meter上检查测量PF、功率精度是否满足要求',
        '预期结果': '4、AC meter上检查测量PF、功率精度满足要求\n有功功率 P 取值（±0.2%，工程设计目标；认证硬指标为0.5%）：\nA相：-13877.7 ~ -13822.2 W\n系统：-13877.7 ~ -13822.2 W\n无功功率 Q 取值（±0.5%，以额定S为基准，SRS §3.5）：\nA相：-24127.4 ~ -23850.4 VAR\n系统：-24127.4 ~ -23850.4 VAR\n视在功率 S 取值（±0.5%，SRS §3.5）：\nA相：27561.5 ~ 27838.5 VA\n系统：27561.5 ~ 27838.5 VA\n功率因数 PF 取值：（Class 0.2，±0.002）\nA相：PF -0.502 ~ -0.498\n系统：PF -0.502 ~ -0.498',
    }, config_path=TEST_CONFIG)
    assert report.passed
