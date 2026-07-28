# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_003_001
模块：IO接入参数设置 → AcuIOM-02 AI设置
标题：AI 页面进入与 Signal Type 保存冒烟

步骤：
  1. 进入 Settings-Devices-AcuIOM-02-点顶部 IO-AI 标签
  2. AI1 Signal Type 选 Voltage(0-10V)，Save
  3. 界面回读并经 Modbus 回读 AI1 Type

预期：成功进入 AcuIOM-02 IO-AI 标签，展示 16 行(AI 1~16)；保存成功；界面回显与
Modbus 回读一致
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_003_001(app_page):
    nav_to_io_ai(app_page)

    assert ai_row_count(app_page) == AI_COUNT, f"AI 标签应展示 {AI_COUNT} 行"

    set_signal_type(app_page, 1, "Voltage(0-10V)")
    saved = save_and_check(app_page)
    assert saved, "AI1 Signal Type=Voltage(0-10V): 保存失败"

    assert get_signal_type_text(app_page, 1) == "Voltage(0-10V)", "AI1 Signal Type 页面回显不符"
    verify_modbus(ai_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AI1 Type")
