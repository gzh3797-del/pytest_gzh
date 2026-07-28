# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_003_008
模块：IO接入参数设置 → AcuIOM-02 AI设置
标题：AI Copy 批量复制 + Reset 恢复出厂（覆盖至 AI16 验 16 通道）

步骤：
  1. 进入 IO-AI 标签
  2. 前置：源 AI1 Signal Type 设为 Current(4-20mA)（区别于默认 Voltage(0-10V)），保存
  3. 勾选源 AI1 → Copy → 勾选目标 AI16 → Apply to Selected → Save
  4. Modbus 回读 AI16 Type，确认被套用为 AI1 的 Current(4-20mA)
  5. 还原：AI16 Reset 出厂；AI1 设回 Voltage(0-10V)

说明（真机行为）：Copy 流程=勾选源→Copy→勾选目标→Apply(先有勾选才可点)→Save；
**源/目标配置相同则保存不提示成功**，故先制造差异并以 Modbus 回读做权威断言；
Reset Selected=对勾选通道恢复出厂设置。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_003_008(app_page):
    nav_to_io_ai(app_page)

    set_signal_type(app_page, 1, "Current(4-20mA)")
    save_and_check(app_page)

    select_row_checkbox(app_page, 1)
    click_copy(app_page, 1)
    step(f"[COPY] 顶部提示: {copied_message_text(app_page)!r}")
    select_row_checkbox(app_page, AI_COUNT)
    click_apply_to_selected(app_page)
    save_and_check(app_page)  # 相同配置不提示成功，以 Modbus 回读为准

    verify_modbus(ai_type_reg(AI_COUNT), SIGNAL_TYPE_ENCODE["Current(4-20mA)"],
                  label=f"AI{AI_COUNT} Type(applied from AI1)")

    select_row_checkbox(app_page, AI_COUNT)
    click_reset_selected(app_page)
    save_and_check(app_page)
    set_signal_type(app_page, 1, "Voltage(0-10V)")
    save_and_check(app_page)
    verify_modbus(ai_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AI1 Type(restore)")
