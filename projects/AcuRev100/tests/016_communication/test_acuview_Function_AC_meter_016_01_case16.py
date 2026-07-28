r"""Function_AC meter_016_01_case16
用例标题: 一台电表通过USB连接到PC，配置不同的slave_id，验证上位机可以正常连接到电表
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Acuview 通过 *USB(COM6)* 连接(Add Connection 行1 ACmeter-USB), 配 SlaveID=1/3,
          上位机 real-time 数值正常。
预期结果: 通过 USB 以不同 slaveID 均连接成功, real-time 显示正常。

2026-07-15 结论: **转手动执行, 不做自动化**(用户确认)。双重阻塞:
  1) 本用例核心是"改 SlaveID(1→3)后 Acuview 以新地址重连", 与 case1/case2 同——此版本 Acuview
     改 SlaveID 后连接必断且不自动重连(需手动关连接窗口重开), 无法自动化逐值验证。端口对调
     (Acuview=USB/COM6 + verify=RS485/COM11 并行)只是并行基建, 解决不了该重连阻塞。
  2) real-time 数值验证依赖测量, 当前板 ADC 损坏为已知 FAIL。
  手动执行: 手动在 Add Connection 选 ACmeter-USB(行1)连接, 改 SlaveID, 关连接重开确认可连。
"""
import pytest

pytestmark = pytest.mark.skip(
    reason="SlaveID改后Acuview不自动重连(需手动关连接重开)+real-time撞ADC已知FAIL, 经用户确认转手动执行")


def test_016_01_case16():
    """转手动执行, 见模块 docstring。"""
