r"""Function_AC meter_004_01_case19
用例标题: CT选型100mA(2000A)，3E4WY，CT一次侧额定电流取5-2000A范围上限2000A，三路显示电流均为1000A，测量Ia,Ib,Ic,Iavg满足精度0.2%
预置条件: 1、Acuview2上位机
2、RTU串口线
3、AcuRev-100电表
4、接线方式3E4WY
测试步骤: 1、连接AcuRev-100电表到Acuview2上，链接状态为connected
2、上位机配置CT Type=100mA、CT Primary=2000（范围上限）
3、源输入三路电流均为10A（经台体CT 20A/100mA，显示系数=Primary/20=100，对应显示1000A），分别配置为对应A相、B相、C相
4、Acuview2上检查Ia、Ib、Ic、Iavg是否均在精度范围内
5、恢复CT Primary=1000（默认值）
预期结果: 4、Acuview2上检查交流电相电流均在精度范围内
Ia：998.0 ~ 1002.0
Ib：998.0 ~ 1002.0
Ic：998.0 ~ 1002.0
Iavg：998.0 ~ 1002.0
5、回读CT Primary=1000，还原成功

生成说明: 方案A——CL3021 设源(0网段 UDP) + USB 口 Modbus 区间断言, 不驱动 Acuview2
GUI(故无锁屏 skipif); 判据/源设定见 case_map.yaml 同编号条目, 复核改 yaml 不改本文件。
⚠ 2026-07-14: CT_PRIMARY@4170 当前固件为枚举索引, 5-2000A连续未落地, case_map 带
needs_review 阻塞闸(运行即拒绝), 待固件支持后解除。
"""
from pathlib import Path

from projects.AcuRev100.tests.helpers_accuracy import run_accuracy_case

PROJECT_ROOT = Path(__file__).resolve().parents[2]      # <模块目录> -> tests -> AcuRev100
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")


def test_004_01_case19():
    report = run_accuracy_case(case_meta={
        '编号': 'Function_AC meter_004_01_case19',
        '标题': 'CT选型100mA(2000A)，3E4WY，CT一次侧额定电流取5-2000A范围上限2000A，三路显示电流均为1000A，测量Ia,Ib,Ic,Iavg满足精度0.2%',
        '预置条件': '1、Acuview2上位机\n2、RTU串口线\n3、AcuRev-100电表\n4、接线方式3E4WY',
        '测试步骤': '1、连接AcuRev-100电表到Acuview2上，链接状态为connected\n2、上位机配置CT Type=100mA、CT Primary=2000（范围上限）\n3、源输入三路电流均为10A（经台体CT 20A/100mA，显示系数=Primary/20=100，对应显示1000A），分别配置为对应A相、B相、C相\n4、Acuview2上检查Ia、Ib、Ic、Iavg是否均在精度范围内\n5、恢复CT Primary=1000（默认值）',
        '预期结果': '4、Acuview2上检查交流电相电流均在精度范围内\nIa：998.0 ~ 1002.0\nIb：998.0 ~ 1002.0\nIc：998.0 ~ 1002.0\nIavg：998.0 ~ 1002.0\n5、回读CT Primary=1000，还原成功',
    }, config_path=TEST_CONFIG)
    assert report.passed
