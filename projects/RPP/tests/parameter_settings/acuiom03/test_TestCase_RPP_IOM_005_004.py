# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_005_004
模块：IO接入参数设置 → AcuIOM-03 DI设置
标题：DI Pulse Constant 非法值校验

步骤：
  1. 进入 IO-DI 标签，DI1 Function=Pulse Counter
  2. Pulse Constant 依次输入：空、abc、@、-1、超上界，逐一 Save

预期：每种非法输入均被阻止保存并提示，配置不写入设备

备注：AcuIOM-03(IOM03P170S04)当前离线，本组用例需设备上线后执行。TODO(需真机
确认) —— 以 save_and_check() 返回 False 作为"被阻止"判据，超上界样例值为占位。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403

_INVALID_VALUES = ["", "abc", "@", "-1", "999999.999"]


def test_TestCase_RPP_IOM_005_004(app_page):
    nav_to_io_di(app_page)
    set_function(app_page, 1, "Pulse Counter")

    for bad_value in _INVALID_VALUES:
        set_pulse_constant(app_page, 1, bad_value)
        saved = save_and_check(app_page)
        assert not saved, f"DI1 Pulse Constant={bad_value!r}: 期望被阻止保存，实际保存成功"
