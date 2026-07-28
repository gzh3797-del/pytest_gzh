# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_006_005
模块：IO接入参数设置 → AcuIOM-04 DI设置
标题：DI Unit 设置

步骤：
  1. 进入 IO-DI 标签，DI1 Function=Pulse Counter
  2. Unit=Cm，Save，回读
  3. 恢复默认并回读确认

预期：保存成功，界面回显 Cm；还原成功

备注：DI1 默认 Unit="Cm" 为探查阶段真机实测原始值（IOM04P193S02 DI1）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403

_DEFAULT_UNIT = "Cm"


def test_TestCase_RPP_IOM_006_005(app_page):
    nav_to_io_di(app_page)
    set_function(app_page, 1, "Pulse Counter")

    set_unit(app_page, 1, "Cm")
    saved = save_and_check(app_page)
    assert saved, "DI1 Unit=Cm: 保存失败"
    assert unit_value(app_page, 1) == "Cm", "DI1 Unit 页面回显不符"

    # 3. 恢复默认（此处默认值本就是 Cm，显式回写确保状态一致）
    set_unit(app_page, 1, _DEFAULT_UNIT)
    saved = save_and_check(app_page)
    assert saved, "DI1 Unit 还原默认: 保存失败"
    assert unit_value(app_page, 1) == _DEFAULT_UNIT, "DI1 Unit 还原后页面回显不符"
