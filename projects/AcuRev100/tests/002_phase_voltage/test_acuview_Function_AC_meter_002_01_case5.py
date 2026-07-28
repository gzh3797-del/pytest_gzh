r"""Function_AC meter_002_01_case5
用例标题: 交流电Van、Vbn、Vcn输入额定电压为(120V,220V,277V),测量Van、Vbn、Vcn、Vlnavg满足精度要求0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的额定电压分别为(120V,220V,277V)。
3、Acuview2上检查交流电相电压是否均在精度范围内，Van范围,Vbn范围,Vcn范围,Vlnavg范围
预期结果: 3、Acuview2上检查交流电相电压均在精度范围内，
Van范围(V):119.760 ~ 120.240，
Vbn范围(V):219.560 ~ 220.440，
Vcn范围(V):276.446 ~ 277.554，
Vlnavg范围(V):205.255 ~ 206.078

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_002_01_case5():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_002_01_case5',
        '标题': '交流电Van、Vbn、Vcn输入额定电压为(120V,220V,277V),测量Van、Vbn、Vcn、Vlnavg满足精度要求0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的额定电压分别为(120V,220V,277V)。\n3、Acuview2上检查交流电相电压是否均在精度范围内，Van范围,Vbn范围,Vcn范围,Vlnavg范围',
        '预期结果': '3、Acuview2上检查交流电相电压均在精度范围内，\nVan范围(V):119.760 ~ 120.240，\nVbn范围(V):219.560 ~ 220.440，\nVcn范围(V):276.446 ~ 277.554，\nVlnavg范围(V):205.255 ~ 206.078',
    }, config_path=TEST_CONFIG)
    assert report.passed
