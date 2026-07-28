r"""Function_AC meter_006_01_case1
用例标题: ABC三相电压角度为(0°  120° 240°)，上位机上检查phaseA、phaseB、PhaseC的相角满足精度±0.5°
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、电源输入ABC三相电压角度为(0°  120° 240°)，电压为100V
3、检查相电压A B C三相的精度是否符合±0.5°
预期结果: 3、检查相电压A B C三相的精度符合±0.5°(判据=设定值±0.5°)

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_006_01_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_006_01_case1',
        '标题': 'ABC三相电压角度为(0°  120° 240°)，上位机上检查phaseA、phaseB、PhaseC的相角满足精度±0.5°',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、电源输入ABC三相电压角度为(0°  120° 240°)，电压为100V\n3、检查相电压A B C三相的精度是否符合±0.5°',
        '预期结果': '3、检查相电压A B C三相的精度符合±0.5°(判据=设定值±0.5°)',
    }, config_path=TEST_CONFIG)
    assert report.passed
