r"""Function_AC meter_002_01_case4
用例标题: 交流电Van、Vbn、Vcn输入电压分别均为0V，1V，9V,电表Van、Vbn、Vcn、Vlnavg无测量值
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、交流电源输入相电压Van、Vbn、Vcn的额定电压为0V。
3、Acuview2上检查交流电相电压是否为0V
4、交流电源输入相电压Van、Vbn、Vcn的额定电压为1V。
5、Acuview2上检查交流电相电压是否为0V，不考虑精度
6、交流电源输入相电压Van、Vbn、Vcn的额定电压为9V。
7、Acuview2上检查交流电相电压是否为0V，不考虑精度
预期结果: 3、5、7、Acuview2上检查交流电相电压为0V

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
2026-07-14 用户裁决(自供电冲突解法): A相常驻最低供电 100V(config
source.supply_guard 可配, 代码 _guard_point 兜底), 0V/1V/9V 只加 Ub/Uc;
判据仅 Vbn/Vcn=0(PRS: 电压<10V 视为不存在), Van/Vlnavg 留手工;
手工步骤8"web页面"为模板残留(AcuRev-100 无 Web), 忽略。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_002_01_case4():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_002_01_case4',
        '标题': '交流电Van、Vbn、Vcn输入电压分别均为0V，1V，9V,电表Van、Vbn、Vcn、Vlnavg无测量值',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、交流电源输入相电压Van、Vbn、Vcn的额定电压为0V。\n3、Acuview2上检查交流电相电压是否为0V\n4、交流电源输入相电压Van、Vbn、Vcn的额定电压为1V。\n5、Acuview2上检查交流电相电压是否为0V，不考虑精度\n6、交流电源输入相电压Van、Vbn、Vcn的额定电压为9V。\n7、Acuview2上检查交流电相电压是否为0V，不考虑精度',
        '预期结果': '3、5、7、Acuview2上检查交流电相电压为0V',
    }, config_path=TEST_CONFIG)
    assert report.passed
