# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_009_003
模块：IO接入参数设置 → AcuIOM-03 DO设置
标题：DO Pulse Width 非法值校验

步骤：
  1. 进入 IO-DO 标签，DO1 Control Mode=Pulse
  2. Pulse Width 依次输入：空、abc、@、-1、超上界，逐一 Save

预期：每种非法输入均被阻止保存并提示，配置不写入设备

备注：AcuIOM-03 离线，需上线后执行（conftest 自动 pytest.skip）。以 save_and_check()
返回 False 作为"被阻止"判据；超上界样例 999999 为占位，请按真机实际量程调整。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_do import *  # noqa: F401, F403

_INVALID_VALUES = ["", "abc", "@", "-1", "999999"]


def test_TestCase_RPP_IOM_009_003(app_page):
    nav_to_io_do(app_page)
    set_control_mode(app_page, 1, "Pulse")

    for bad_value in _INVALID_VALUES:
        set_pulse_width(app_page, 1, bad_value)
        saved = save_and_check(app_page)
        assert not saved, f"DO1 Pulse Width={bad_value!r}: 期望被阻止保存，实际保存成功"
