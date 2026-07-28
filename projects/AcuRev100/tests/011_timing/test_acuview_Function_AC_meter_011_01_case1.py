r"""Function_AC meter_011_01_case1
用例标题: 铅封未封闭，通过上位机设置时间，年/月/日和时/分/秒，遍历所有时间，遍历所有日期(包括闰年)，遍历所有周
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 2、遍历写入9组时间(边界2000/2099、周一~周日、闰年2024-02-29)
3、逐组读回校验
(掉电重启步骤由 allow_power_cycle 把关)
预期结果: 每组时间设置成功且读回一致(总秒差≤10s); 星期显示MANUAL; 还原PC时间

生成说明: 计时类(方案A)——Modbus 直连读写时钟/计数器寄存器, 需要负载条件的段由 CL3021
控源; 不驱动 Acuview2 GUI。判据/段定义见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
掉电重启类步骤由 config run.allow_power_cycle 把关(默认 false 记 MANUAL)。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_timing import run_timing_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_011_01_case1():
    report = run_timing_case(case_meta={
        '编号': 'Function_AC meter_011_01_case1',
        '标题': '铅封未封闭，通过上位机设置时间，年/月/日和时/分/秒，遍历所有时间，遍历所有日期(包括闰年)，遍历所有周',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '2、遍历写入9组时间(边界2000/2099、周一~周日、闰年2024-02-29)\n3、逐组读回校验\n(掉电重启步骤由 allow_power_cycle 把关)',
        '预期结果': '每组时间设置成功且读回一致(总秒差≤10s); 星期显示MANUAL; 还原PC时间',
    }, config_path=TEST_CONFIG)
    assert report.passed
