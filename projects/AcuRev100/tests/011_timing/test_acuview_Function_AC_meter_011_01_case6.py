r"""Function_AC meter_011_01_case6
用例标题: 使用100mA（200A）的CT，设备有负载的累计时间；即电流输入 > Ist (0.2A) 且电压 > 10V
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、Ist以上点(0.3A，>Ist 0.2A)等待后校验负载时间累计
3、1A点同上
4、低于Ist/零流点校验不累计
预期结果: 有负载(>Ist)段 Δ≈等待时长; 无负载段 Δ≈0

生成说明: 计时类(方案A)——Modbus 直连读写时钟/计数器寄存器, 需要负载条件的段由 CL3021
控源; 不驱动 Acuview2 GUI。判据/段定义见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
掉电重启类步骤由 config run.allow_power_cycle 把关(默认 false 记 MANUAL)。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_timing import run_timing_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_011_01_case6():
    report = run_timing_case(case_meta={
        '编号': 'Function_AC meter_011_01_case6',
        '标题': '使用100mA（200A）的CT，设备有负载的累计时间；即电流输入 > Ist (0.2A) 且电压 > 10V',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、Ist以上点(0.3A，>Ist 0.2A)等待后校验负载时间累计\n3、1A点同上\n4、低于Ist/零流点校验不累计',
        '预期结果': '有负载(>Ist)段 Δ≈等待时长; 无负载段 Δ≈0',
    }, config_path=TEST_CONFIG)
    assert report.passed
