r"""Function_AC meter_010_13_case3
用例标题: 3E4WY V-N 反接阈值 K=1.3 边界
预置条件: 1、Acuview2上位机（RS-485，Modbus RTU）
2、可程控三相功率源
3、AcuRev-100电表
4、接线检查功能已启用（默认启用）
5、Dip Switch 处于解锁状态（允许接线配置写入）
测试步骤: 1. 配置接线方式 = 3E4WY，相序配置 = ABC（Modbus 寄存器 0x1063 Phase Order = 0）
2. 程控源输出正常基准：ua=ub=uc=220V，qua=0°/qub=240°/quc=120°，ia=ib=ic=5A，qia=0°/qib=240°/qic=120°，freq=50Hz；关→开一次接线检查开关触发立即检测（开关开启瞬间检查一次；常开时检测周期 1 次/min），回读接线检查结果确认全部正常
3. 修改程控源输出（刚触发侧）：ua=150V、ub=uc=195V（Vbn/Van=Vcn/Van=1.30 刚达阈值）；关→开接线检查开关触发立即检测（或等待 ≥60s）
4. 经 Acuview2 Reading—接线检查 页与 Modbus 回读接线检查结果
5. 修改程控源输出（刚不触发侧）：ua=150V、ub=uc=190V（比值 1.267<1.3）；触发立即检测后回读接线检查结果
6. 程控源还原步骤 2 正常基准值；触发立即检测后回读确认告警全部清除
预期结果: 2. 接线检查结果全部正常
4. 上报 电压端 Va-Vn 反接（条件8）
5. 无反接告警，无任何接线告警；三相不对称>10% 致相序算法 Step2 失败，voltagePhaseOrder=2，但输出 2 不参与条件13 比较，不报相序配置错误（已裁定 2026-07-14）
6. 告警清除，接线检查结果恢复全部正常

生成说明: 方案A——CL3021 设源(故障注入) + 开关重开触发立即检测 + USB 口 Modbus 错误码位断言,
不驱动 Acuview2 GUI(故无锁屏 skipif); 期望位值/源设定见 case_map_wiring.yaml 同编号条目,
复核改 yaml 不改本文件。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_wiring import run_wiring_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # wiring_check -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_010_13_case3():
    report = run_wiring_case(case_meta={
        '编号': 'Function_AC meter_010_13_case3',
        '标题': '3E4WY V-N 反接阈值 K=1.3 边界',
        '预置条件': '1、Acuview2上位机（RS-485，Modbus RTU）\n2、可程控三相功率源\n3、AcuRev-100电表\n4、接线检查功能已启用（默认启用）\n5、Dip Switch 处于解锁状态（允许接线配置写入）',
        '测试步骤': '1. 配置接线方式 = 3E4WY，相序配置 = ABC（Modbus 寄存器 0x1063 Phase Order = 0）\n2. 程控源输出正常基准：ua=ub=uc=220V，qua=0°/qub=240°/quc=120°，ia=ib=ic=5A，qia=0°/qib=240°/qic=120°，freq=50Hz；关→开一次接线检查开关触发立即检测（开关开启瞬间检查一次；常开时检测周期 1 次/min），回读接线检查结果确认全部正常\n3. 修改程控源输出（刚触发侧）：ua=150V、ub=uc=195V（Vbn/Van=Vcn/Van=1.30 刚达阈值）；关→开接线检查开关触发立即检测（或等待 ≥60s）\n4. 经 Acuview2 Reading—接线检查 页与 Modbus 回读接线检查结果\n5. 修改程控源输出（刚不触发侧）：ua=150V、ub=uc=190V（比值 1.267<1.3）；触发立即检测后回读接线检查结果\n6. 程控源还原步骤 2 正常基准值；触发立即检测后回读确认告警全部清除',
        '预期结果': '2. 接线检查结果全部正常\n4. 上报 电压端 Va-Vn 反接（条件8）\n5. 无反接告警，无任何接线告警；三相不对称>10% 致相序算法 Step2 失败，voltagePhaseOrder=2，但输出 2 不参与条件13 比较，不报相序配置错误（已裁定 2026-07-14）\n6. 告警清除，接线检查结果恢复全部正常',
    }, config_path=TEST_CONFIG)
    assert report.passed
