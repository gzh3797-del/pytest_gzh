# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_001_001
模块：IO接入参数设置 → AcuIOM-01 AI设置
标题：AI 页面进入与 Signal Type 保存冒烟

步骤：
  1. 进入 Settings-Devices-AcuIOM-01-点顶部 IO-AI 标签
  2. AI1 Signal Type 下拉选 Voltage(0-10V)，Save
  3. 界面回读 AI1 Signal Type；经 Modbus FC03 回读 AI1 Type 寄存器(12288)

预期：
  1. 成功进入 AcuIOM-01 IO-AI 标签，展示 8 行(AI 1~8)
  2. 保存成功，无报错
  3. 界面回显 Voltage(0-10V)；Modbus 回读 AI1 Type 编码值与所选一致
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_001_001(app_page):
    nav_to_io_ai(app_page)

    # 1. 页面成功展示 8 行 (AI 1~8)
    assert ai_row_count(app_page) == AI_COUNT, f"AI 标签应展示 {AI_COUNT} 行"

    # 2. AI1 Signal Type -> Voltage(0-10V)，Save
    set_signal_type(app_page, 1, "Voltage(0-10V)")
    saved = save_and_check(app_page)
    assert saved, "AI1 Signal Type=Voltage(0-10V): 保存失败"

    # 3. 页面回显 + Modbus 回读一致
    assert get_signal_type_text(app_page, 1) == "Voltage(0-10V)", "AI1 Signal Type 页面回显不符"
    verify_modbus(ai_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AI1 Type")
