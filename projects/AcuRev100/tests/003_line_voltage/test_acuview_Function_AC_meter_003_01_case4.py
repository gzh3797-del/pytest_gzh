r"""Function_AC meter_003_01_case4
用例标题: 交流电Van、Vbn、Vcn输入额定电压为(120V,220V,270V), θA、θB、θC分别为(0°、60°、120°)测量Vab、Vbc、Vca、Vllavg满足精度要求0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的电压为(120V,220V,270V), θA、θB、θC分别为(0°、60°、120°)。
3、Acuview2上检查交流电线电压是否均在精度范围内，
预期结果: 3、Acuview2上检查交流电线电压均在精度范围内
Vab范围(V)：190.406 ~ 191.170
Vca范围(V)：345.285 ~ 346.669
Vbc范围(V)：248.300 ~ 249.295
Vllavg范围(V)：261.330 ~ 262.378

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_003_01_case4():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_003_01_case4',
        '标题': '交流电Van、Vbn、Vcn输入额定电压为(120V,220V,270V), θA、θB、θC分别为(0°、60°、120°)测量Vab、Vbc、Vca、Vllavg满足精度要求0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的电压为(120V,220V,270V), θA、θB、θC分别为(0°、60°、120°)。\n3、Acuview2上检查交流电线电压是否均在精度范围内，',
        '预期结果': '3、Acuview2上检查交流电线电压均在精度范围内\nVab范围(V)：190.406 ~ 191.170\nVca范围(V)：345.285 ~ 346.669\nVbc范围(V)：248.300 ~ 249.295\nVllavg范围(V)：261.330 ~ 262.378',
    }, config_path=TEST_CONFIG)
    assert report.passed
