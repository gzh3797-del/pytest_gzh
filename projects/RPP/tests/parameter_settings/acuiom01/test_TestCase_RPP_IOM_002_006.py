# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_002_006
模块：IO接入参数设置 → AcuIOM-01 AO设置
标题：AO Copy 批量复制 + Reset 恢复出厂

步骤：
  1. 进入 IO-AO 标签
  2. 前置：把源 AO1 Signal Type 设为 Current(4-20mA)（区别于默认 Voltage(0-10V)），保存
  3. 勾选源 AO1 → 点 Copy → 勾选目标 AO2 → 点 Apply to Selected → Save
  4. Modbus 回读 AO2 Type，确认已被套用为 AO1 的 Current(4-20mA)
  5. 还原：勾选 AO2 → Reset Selected 恢复出厂；AO1 设回 Voltage(0-10V)

预期：
  3. Apply 后 AO2 被套用为 AO1 配置
  4. Modbus 回读 AO2 Type = 3（Current(4-20mA)）
  5. AO2 恢复出厂默认、AO1 恢复 Voltage(0-10V)

说明（真机行为）：
  - Copy 正确流程（用户确认）：勾选源 → Copy → 勾选目标 → Apply to Selected（先有勾选才可点）→ Save。
  - **若源/目标通道配置本就相同，保存不提示"修改成功"**；故本用例先把源 AO1 改成非默认值制造差异，
    并以 Modbus 回读做权威断言，不依赖保存成功提示（save_and_check 不作硬断言）。
  - Reset Selected = 对勾选通道恢复出厂设置。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_002_006(app_page):
    nav_to_io_ao(app_page)

    # 2. 前置：源 AO1 设为非默认 Signal Type，保证 Apply 到 AO2 确有变化
    set_signal_type(app_page, 1, "Current(4-20mA)")
    save_and_check(app_page)

    # 3. 勾选源 AO1 → Copy → 勾选目标 AO2 → Apply to Selected → Save
    select_row_checkbox(app_page, 1)
    click_copy(app_page, 1)
    step(f"[COPY] 顶部提示: {copied_message_text(app_page)!r}")
    select_row_checkbox(app_page, 2)
    click_apply_to_selected(app_page)
    save_and_check(app_page)  # 相同配置时不提示成功，故不硬断言，以 Modbus 回读为准

    # 4. Modbus 回读 AO2 Type = Current(4-20mA)=3，确认被套用为 AO1
    verify_modbus(ao_type_reg(2), SIGNAL_TYPE_ENCODE["Current(4-20mA)"],
                  label="AO2 Type(applied from AO1)")

    # 5. 还原：AO2 Reset 出厂；AO1 设回默认 Voltage(0-10V)
    select_row_checkbox(app_page, 2)
    click_reset_selected(app_page)
    save_and_check(app_page)
    set_signal_type(app_page, 1, "Voltage(0-10V)")
    save_and_check(app_page)
    verify_modbus(ao_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AO1 Type(restore)")
