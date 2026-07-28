# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_005_005
模块：IO接入参数设置 → AcuIOM-03 DI设置
标题：DI Unit 设置

步骤：
  1. 进入 IO-DI 标签，DI1 Function=Pulse Counter
  2. Unit=Cm，Save，回读
  3. 恢复默认并回读确认

预期：保存成功，界面回显 Cm；还原成功

备注：AcuIOM-03(IOM03P170S04)当前离线，本组用例需设备上线后执行。TODO(需真机
确认) —— DI1 出厂默认 Unit 具体值未知（该设备探查阶段全程离线），本用例改用
"运行期先捕获当前值、用后再还原"策略。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_005_005(app_page):
    nav_to_io_di(app_page)
    set_function(app_page, 1, "Pulse Counter")

    original_unit = unit_value(app_page, 1)

    set_unit(app_page, 1, "Cm")
    saved = save_and_check(app_page)
    assert saved, "DI1 Unit=Cm: 保存失败"
    assert unit_value(app_page, 1) == "Cm", "DI1 Unit 页面回显不符"

    # 3. 用运行期捕获的原始值还原
    set_unit(app_page, 1, original_unit)
    saved = save_and_check(app_page)
    assert saved, "DI1 Unit 还原原始值: 保存失败"
    assert unit_value(app_page, 1) == original_unit, "DI1 Unit 还原后页面回显不符"
