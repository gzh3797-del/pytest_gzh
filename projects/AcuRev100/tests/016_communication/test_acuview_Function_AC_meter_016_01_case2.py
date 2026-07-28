r"""Function_AC meter_016_01_case2
用例标题: 上位机上通过Rtu/USB，配置slaveID为0、255失败
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: Communication 页尝试把 Modbus Slave ID 设为 0、255(越界, 合法 1-247), 逐个下发;
          每次经寄存器(Modbus)回读确认 SlaveID 未变为该越界值(被拒/被钳制)。
预期结果: 0/255 均无法写入生效, 寄存器回读保持合法值不变; 还原=1。

2026-07-15 结论: **转手动执行, 不做自动化**(与 case1 同因)。
  越界值若被上位机接受/截断而真正改变了 SlaveID, 会导致 Acuview 断连且不自动重连(需手动关连接
  窗口重开)。历史上该越界写入曾把电表 SlaveID 改成 25 引发断链。判据需人工目视上位机拒绝提示 +
  寄存器回读确认地址未变, 经用户确认转手动执行。
"""
import pytest

pytestmark = pytest.mark.skip(
    reason="SlaveID改后Acuview连接必断且不自动重连(需手动关连接窗口重开), 经用户确认转手动执行")


def test_016_01_case2():
    """转手动执行, 见模块 docstring。"""
