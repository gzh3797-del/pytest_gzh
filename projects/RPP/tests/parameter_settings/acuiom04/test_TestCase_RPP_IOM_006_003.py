# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_006_003
模块：IO接入参数设置 → AcuIOM-04 DI设置
标题：DI Pulse Constant 有效值下发回读

步骤：
  1. 进入 IO-DI 标签，DI1 Function=Pulse Counter
  2. Pulse Constant=100.000，Save
  3. 界面回读；Modbus 回读 DI1 Pulse Constant 配置寄存器
  4. 恢复 Function=Status Monitor 回读确认

预期：保存成功；界面回显 100.000，Modbus 回读一致；还原成功

备注：DI1 默认 Pulse Constant=100.000、Unit="Cm"、Function=Status Monitor 为探查
阶段真机实测原始值（IOM04P193S02 DI1，disabled 态仍保留原值可见）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_006_003(app_page):
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
