# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_005_003
模块：IO接入参数设置 → AcuIOM-03 DI设置
标题：DI Pulse Constant 有效值下发回读

步骤：
  1. 进入 IO-DI 标签，DI1 Function=Pulse Counter
  2. Pulse Constant=100.000，Save
  3. 界面回读；经 Modbus 回读 DI1 Pulse Constant 配置寄存器
  4. 恢复 Function=Status Monitor 回读确认

预期：保存成功；界面回显 100.000，Modbus 回读一致；还原成功

备注：AcuIOM-03(IOM03P170S04)当前离线，本组用例需设备上线后执行；DI1 出厂默认
Pulse Constant/Unit 具体值未经真机确认（该设备探查阶段全程离线，无法读取），
本用例仅还原 Function=Status Monitor，不假定 Pulse Constant 具体出厂值。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_005_003(app_page):
    nav_to_io_di(app_page)
    set_function(app_page, 1, "Pulse Counter")

    set_pulse_constant(app_page, 1, "100.000")
    saved = save_and_check(app_page)
    assert saved, "DI1 Pulse Constant=100.000: 保存失败"

    verify_modbus_scaled(di_pulse_constant_reg(1), 100.000, label="DI1 Pulse Constant")

    # 4. 恢复默认 Function=Status Monitor
    set_function(app_page, 1, "Status Monitor")
    saved = save_and_check(app_page)
    assert saved, "DI1 还原 Function=Status Monitor: 保存失败"
    verify_modbus(di_type_reg(1), FUNCTION_ENCODE["Status Monitor"], label="DI1 Type(restore)")
