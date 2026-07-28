r"""Function_AC meter_002_01_case6
用例标题: 交流电Van、Vbn、Vcn输入额定电压为(120V,120V,120V)顺序，且φA、φB、φC分别为(0°、130°、250°)，Van、Vbn、Vcn、Vlnavg测量值满足精度要求0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的额定电压为120V。
3、电表相序设置为ABC，φA、φB、φC分别为(0°、130°、250°)
4、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值是否满足精度要求
预期结果: 4、Acuview2上检查交流电相电压均在精度范围内，范围(V):119.760 ~ 120.240，
Vlnavg范围(V)：119.760 ~ 120.240

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_002_01_case6():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_002_01_case6',
        '标题': '交流电Van、Vbn、Vcn输入额定电压为(120V,120V,120V)顺序，且φA、φB、φC分别为(0°、130°、250°)，Van、Vbn、Vcn、Vlnavg测量值满足精度要求0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的额定电压为120V。\n3、电表相序设置为ABC，φA、φB、φC分别为(0°、130°、250°)\n4、Acuview2上检查交流电相电压Van、Vbn、Vcn、Vlnavg测量值是否满足精度要求',
        '预期结果': '4、Acuview2上检查交流电相电压均在精度范围内，范围(V):119.760 ~ 120.240，\nVlnavg范围(V)：119.760 ~ 120.240',
    }, config_path=TEST_CONFIG)
    assert report.passed
