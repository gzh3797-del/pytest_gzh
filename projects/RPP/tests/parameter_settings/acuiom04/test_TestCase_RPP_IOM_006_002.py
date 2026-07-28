# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_006_002
模块：IO接入参数设置 → AcuIOM-04 DI设置
标题：DI Function 切换与 Pulse Constant/Unit 联动

步骤：
  1. 进入 IO-DI 标签
  2. DI1 Function 选 Status Monitor，观察 Pulse Constant、Unit 状态
  3. DI1 切 Pulse Counter，再观察
  4. Save 后回读 DI1 Function

预期：
  2. Status Monitor 下 Pulse Constant、Unit 禁用
  3. Pulse Counter 下两字段可编辑
  4. 保存成功，回显 Pulse Counter
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_006_002(app_page):
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
