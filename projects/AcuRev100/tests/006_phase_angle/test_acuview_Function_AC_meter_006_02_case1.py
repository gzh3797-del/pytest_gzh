r"""Function_AC meter_006_02_case1
用例标题: ABC三相电流相角分设置为0°、1°、5°，上位机上检查input1的相角满足精度±0.5°
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、电源输入input1~3的电流相角为(0°、1°、5°)，电流为5A
3、检查input1~3的电流相角的精度是否符合±0.5°
预期结果: 3、检查input1~3的电流相角的精度符合±0.5°(判据=设定值±0.5°)
(自动化补充判据, 2026-07-28) 电流幅值 Ia/Ib/Ic = 49.9 ~ 50.1 A
  (源 5A × CT Primary 200 ÷ 台体CT 20A = 50A, ±0.2%): 手工用例只写了相角, 台面某相电流
  回路不通时相角读数即为噪声, 用例只会报"角度不对"而不暴露"电流根本没进表"。手工 xlsx 未改。

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。

⚠️ 已知设备侧问题(2026-07-28 实测, 当前固件非正式版本, 暂不处理; 006_02_case2 同症状):
    **电流相角寄存器 PHASE_A/B/C_CURRENT_PHASE_ANGLE 读数与真实相位不符, 且跨运行不可复现**,
    故本用例的 Ang_IA/IB/IC 会稳定 FAIL —— 判据没错, 是表读错。证据(case2 输入
    qU=0/240/120、qI=60/90/120、5A):
      · PF 三相 = 0.534 / -0.883 / 0.999, 与命令的逐相角差 60°/210°/0° 对得上(反推角
        57.7°/152.0°/2.6°, 误差≤2.6°) ⇒ 源确实逐相施加了电流角, 表也算对了相位关系;
      · θU 三相 = 0 / 240.24 / 120.0 ⇒ 电压角寄存器正常;
      · 而 θI 同一输入两次运行分别读 45/105/75.9 与 315/195/75(真值 60/90/120), 两次都错、
        且互不一致 ⇒ 不是固定偏移, 也不是判据口径问题(绝对角/相对角都套不上)。
    同批还发现表的**幅值读数整体异常**: 电压读 15249V(命令 100V, ≈152 倍)、电流读 39690A
    (期望 50A, ≈794 倍) → 功率寄存器因此饱和在 214748.36。上位机 Acuview2 上看电压同样不对
    (2026-07-28 用户确认) ⇒ 表侧测量链路整体问题, 唯 PF / 频率 / θU 正常。
    ⇒ 判据/寄存器映射均无需改动; 待固件转正式版本后整体重测, 仍复现则按缺陷提研发。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_006_02_case1():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_006_02_case1',
        '标题': 'ABC三相电流相角分设置为0°、1°、5°，上位机上检查input1的相角满足精度±0.5°',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、电源输入input1~3的电流相角为(0°、1°、5°)，电流为5A\n3、检查input1~3的电流相角的精度是否符合±0.5°',
        '预期结果': '3、检查input1~3的电流相角的精度符合±0.5°(判据=设定值±0.5°)\n'
                    '(自动化补充判据) 电流幅值 Ia/Ib/Ic = 49.9 ~ 50.1 A'
                    '(源5A × CT Primary 200 ÷ 台体CT 20A = 50A, ±0.2%)',
    }, config_path=TEST_CONFIG)
    assert report.passed
