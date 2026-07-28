# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_003_002
模块：IO接入参数设置 → AcuIOM-02 AI设置
标题：AI Signal Type 四种输入类型遍历

步骤：
  1. 进入 IO-AI 标签
  2. AI1 Signal Type 依次设四种类型，每次 Save
  3. 每次 Modbus 回读 AI1 Type
  4. 恢复 Voltage(0-10V) 回读确认

预期：四种均保存成功；Modbus 回读与所选一致；还原成功
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403

_TYPES = ["Voltage(0-10V)", "Voltage(2-10V)", "Current(0-20mA)", "Current(4-20mA)"]


def test_TestCase_RPP_IOM_003_002(app_page):
    nav_to_io_ai(app_page)

    for option in _TYPES:
        set_signal_type(app_page, 1, option)
        saved = save_and_check(app_page)
        assert saved, f"AI1 Signal Type={option}: 保存失败"
        verify_modbus(ai_type_reg(1), SIGNAL_TYPE_ENCODE[option], label=f"AI1 Type({option})")

    set_signal_type(app_page, 1, "Voltage(0-10V)")
    saved = save_and_check(app_page)
    assert saved, "AI1 还原 Voltage(0-10V): 保存失败"
    verify_modbus(ai_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AI1 Type(restore)")
