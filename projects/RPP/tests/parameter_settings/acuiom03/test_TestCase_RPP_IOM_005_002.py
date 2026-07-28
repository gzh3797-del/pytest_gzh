# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_005_002
模块：IO接入参数设置 → AcuIOM-03 DI设置
标题：DI Function 切换与 Pulse Constant/Unit 联动

步骤：
  1. 进入 IO-DI 标签
  2. DI1 Function 选 Status Monitor，观察 Pulse Constant、Unit 状态
  3. DI1 切 Pulse Counter，再观察
  4. Save 后回读 DI1 Function

预期：
  2. Status Monitor 下 Pulse Constant、Unit 禁用(灰态)
  3. Pulse Counter 下两字段可编辑
  4. 保存成功，回显 Pulse Counter

备注：AcuIOM-03(IOM03P170S04)当前离线，本组用例需设备上线后执行；DO/RO 标签
结构待上线后核实补充。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_005_002(app_page):
    nav_to_io_di(app_page)

    set_function(app_page, 1, "Status Monitor")
    assert pulse_constant_disabled(app_page, 1), "Status Monitor 下 Pulse Constant 应禁用"
    assert unit_disabled(app_page, 1), "Status Monitor 下 Unit 应禁用"

    set_function(app_page, 1, "Pulse Counter")
    assert not pulse_constant_disabled(app_page, 1), "Pulse Counter 下 Pulse Constant 应可编辑"
    assert not unit_disabled(app_page, 1), "Pulse Counter 下 Unit 应可编辑"

    saved = save_and_check(app_page)
    assert saved, "DI1 Function=Pulse Counter: 保存失败"
    assert get_function_text(app_page, 1) == "Pulse Counter", "DI1 Function 页面回显不符"
