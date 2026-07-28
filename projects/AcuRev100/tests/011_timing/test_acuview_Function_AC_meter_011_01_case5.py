r"""Function_AC meter_011_01_case5
用例标题: 设备开始运行以来的累计时间正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 1、读运行时间T1
2、等待后再读T2, 校验T2-T1≈等待时长
(掉电保持步骤由 allow_power_cycle 把关)
预期结果: 运行时间随真实时间累计(Δ≈等待时长±10s)

生成说明: 计时类(方案A)——Modbus 直连读写时钟/计数器寄存器, 需要负载条件的段由 CL3021
控源; 不驱动 Acuview2 GUI。判据/段定义见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
掉电重启类步骤由 config run.allow_power_cycle 把关(默认 false 记 MANUAL)。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_timing import run_timing_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_011_01_case5():
    report = run_timing_case(case_meta={
        '编号': 'Function_AC meter_011_01_case5',
        '标题': '设备开始运行以来的累计时间正常',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、读运行时间T1\n2、等待后再读T2, 校验T2-T1≈等待时长\n(掉电保持步骤由 allow_power_cycle 把关)',
        '预期结果': '运行时间随真实时间累计(Δ≈等待时长±10s)',
    }, config_path=TEST_CONFIG)
    assert report.passed
