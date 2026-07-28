r"""Function_AC meter_003_01_case2
用例标题: 交流电Van、Vbn、Vcn均输入临界电压10.2V、20V,电表有测量值，测量Vab、Vbc、Vca、Vllavg精度没有要求
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的额定电压为10.2V。
3、Acuview2上检查交流电线电压是否有测量显示， 精度没有要求
4、交流电源输入相电压Van、Vbn、Vcn的额定电压为20V。
5、Acuview2上检查交流电线电压是否有测量显示, 精度没有要求
预期结果: 3、Acuview2上检查交流电线电压有测量显示，精度没有要求
5、Acuview2上检查交流电线电压有测量显示，精度没有要求

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
2026-07-14 用户裁决(自供电冲突解法): A相常驻最低供电 100V(config
source.supply_guard 可配, 代码 _guard_point 兜底), 10.2V/20V 只加 Ub/Uc;
线电压判据仅 Vbc=U×√3±10%(10.2V→17.67V, 20V→34.64V), Vab/Vca/Vllavg 含A相留手工。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_003_01_case2():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_003_01_case2',
        '标题': '交流电Van、Vbn、Vcn均输入临界电压10.2V、20V,电表有测量值，测量Vab、Vbc、Vca、Vllavg精度没有要求',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的额定电压为10.2V。\n3、Acuview2上检查交流电线电压是否有测量显示， 精度没有要求\n4、交流电源输入相电压Van、Vbn、Vcn的额定电压为20V。\n5、Acuview2上检查交流电线电压是否有测量显示, 精度没有要求',
        '预期结果': '3、Acuview2上检查交流电线电压有测量显示，精度没有要求\n5、Acuview2上检查交流电线电压有测量显示，精度没有要求',
    }, config_path=TEST_CONFIG)
    assert report.passed
