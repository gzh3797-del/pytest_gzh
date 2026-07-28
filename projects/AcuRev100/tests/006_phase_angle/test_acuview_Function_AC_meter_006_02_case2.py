r"""Function_AC meter_006_02_case2
用例标题: ABC三相电流相角分设置为60°、90°、120°，上位机上检查input1的相角满足精度±0.5°
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、电源输入input1~3的电流相角为(60°、90°、120°)，电流为5A
3、检查input1~3的电流相角的精度是否符合±0.5°
预期结果: 3、检查input1~3的电流相角的精度符合±0.5°(判据=设定值±0.5°)
(自动化补充判据, 2026-07-28) 电流幅值 Ia/Ib/Ic = 49.9 ~ 50.1 A
  (源 5A × CT Primary 200 ÷ 台体CT 20A = 50A, ±0.2%): 手工用例只写了相角, 台面某相电流
  回路不通时相角读数即为噪声, 用例只会报"角度不对"而不暴露"电流根本没进表"。手工 xlsx 未改。

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_006_02_case2():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_006_02_case2',
        '标题': 'ABC三相电流相角分设置为60°、90°、120°，上位机上检查input1的相角满足精度±0.5°',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、电源输入input1~3的电流相角为(60°、90°、120°)，电流为5A\n3、检查input1~3的电流相角的精度是否符合±0.5°',
        '预期结果': '3、检查input1~3的电流相角的精度符合±0.5°(判据=设定值±0.5°)\n'
                    '(自动化补充判据) 电流幅值 Ia/Ib/Ic = 49.9 ~ 50.1 A'
                    '(源5A × CT Primary 200 ÷ 台体CT 20A = 50A, ±0.2%)',
    }, config_path=TEST_CONFIG)
    assert report.passed
