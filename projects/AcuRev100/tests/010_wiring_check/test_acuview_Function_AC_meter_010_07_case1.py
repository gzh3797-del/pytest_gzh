r"""Function_AC meter_010_07_case1
用例标题: 1E2W 正常接线基准：无告警且不检测相序
预置条件: 1、Acuview2上位机（RS-485，Modbus RTU）
2、可程控三相功率源
3、AcuRev-100电表
4、接线检查功能已启用（默认启用）
5、Dip Switch 处于解锁状态（允许接线配置写入）
测试步骤: 1. 配置接线方式 = 1E2W（经 Acuview2 Setting 或 Modbus 下发）
2. 程控源输出正常基准：ua=220V/qua=0°，ia=5A/qia=0°，ub=uc=0V，ib=ic=0A，freq=50Hz；关→开一次接线检查开关触发立即检测（开关开启瞬间检查一次；常开时检测周期 1 次/min）
3. 经 Acuview2 Reading—接线检查 页与 Modbus 回读接线检查结果与 voltagePhaseOrder
预期结果: 3. 接线检查结果全部正常，无任何告警；voltagePhaseOrder = 2（1E2W 不检测相序）；LED 绿色（MANUAL 目视）

生成说明: 方案A——CL3021 设源(故障注入) + 开关重开触发立即检测 + USB 口 Modbus 错误码位断言,
不驱动 Acuview2 GUI(故无锁屏 skipif); 期望位值/源设定见 case_map_wiring.yaml 同编号条目,
复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_wiring import run_wiring_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # wiring_check -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_010_07_case1():
    report = run_wiring_case(case_meta={
        '编号': 'Function_AC meter_010_07_case1',
        '标题': '1E2W 正常接线基准：无告警且不检测相序',
        '预置条件': '1、Acuview2上位机（RS-485，Modbus RTU）\n2、可程控三相功率源\n3、AcuRev-100电表\n4、接线检查功能已启用（默认启用）\n5、Dip Switch 处于解锁状态（允许接线配置写入）',
        '测试步骤': '1. 配置接线方式 = 1E2W（经 Acuview2 Setting 或 Modbus 下发）\n2. 程控源输出正常基准：ua=220V/qua=0°，ia=5A/qia=0°，ub=uc=0V，ib=ic=0A，freq=50Hz；关→开一次接线检查开关触发立即检测（开关开启瞬间检查一次；常开时检测周期 1 次/min）\n3. 经 Acuview2 Reading—接线检查 页与 Modbus 回读接线检查结果与 voltagePhaseOrder',
        '预期结果': '3. 接线检查结果全部正常，无任何告警；voltagePhaseOrder = 2（1E2W 不检测相序）；LED 绿色（MANUAL 目视）',
    }, config_path=TEST_CONFIG)
    assert report.passed
