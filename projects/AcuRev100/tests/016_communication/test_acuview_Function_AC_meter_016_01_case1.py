r"""Function_AC meter_016_01_case1
用例标题: 上位机上通过Rtu/USB，配置slaveID为5、247，上位机上显示正常
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Acuview 在 Communication 页把 Modbus Slave ID 依次设 5、247 下发; 每次经寄存器(Modbus)
          回读确认新地址生效, 再关闭连接窗口重新打开使上位机以新/默认地址重连显示正常。
预期结果: SlaveID 5/247 写入成功, 寄存器回读一致, 上位机重连后显示正常。

2026-07-15 结论: **转手动执行, 不做自动化**。
  实测此版本 Acuview 改 SlaveID 后连接立即断开(报 "Connect Failed!"), 且不会自动重连——
  Connection→Connect 亦无效, 必须手动关闭该连接窗口再重新打开(重新 Add Connection)才能重连。
  逐值(5→还原→247)自动化需在每值间自动关/开连接窗口, 流程脆弱且价值低, 经用户确认转手动。
  (电表侧对合法 SlaveID 的接受能力已由自动化验证过: 设 5 后寄存器回读=5 通过。)
"""
import pytest

pytestmark = pytest.mark.skip(
    reason="SlaveID改后Acuview连接必断且不自动重连(需手动关连接窗口重开), 经用户确认转手动执行")


def test_016_01_case1():
    """转手动执行, 见模块 docstring。"""
