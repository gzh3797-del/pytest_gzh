r"""Function_AC meter_004_02_case1
用例标题: CT选型333mV(200A)，1E2W输入InputA电流为100A，测量InputA电流精度0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、1E2W
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、1E2W输入InputA电流为2.5A
3、Acuview2上检查电流InputA是否均在精度范围内
预期结果: 3、Acuview2上检查电流InputA在精度范围内
IA:	99.800	~	100.200

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_02_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_02_case1',
        '标题': 'CT选型333mV(200A)，1E2W输入InputA电流为100A，测量InputA电流精度0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、1E2W',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、1E2W输入InputA电流为2.5A\n3、Acuview2上检查电流InputA是否均在精度范围内',
        '预期结果': '3、Acuview2上检查电流InputA在精度范围内\nIA:\t99.800\t~\t100.200',
    }, config_path=TEST_CONFIG)
    assert report.passed
