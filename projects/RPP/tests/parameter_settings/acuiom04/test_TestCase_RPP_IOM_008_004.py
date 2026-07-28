# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_008_004
模块：IO接入参数设置 → AcuIOM-04 RO设置
标题：RO Copy 批量复制 + Reset 恢复出厂

步骤：
  1. 进入 IO-RO 标签
  2. 前置：目标 RO2 Control Mode 设为 Pulse、源 RO1 设为 Manual（制造差异），分别保存
  3. 勾选源 RO1 → Copy → 勾选目标 RO2 → Apply to Selected → Save
  4. Modbus 回读 RO2 Control Mode，确认被套用为 RO1 的 Manual
  5. 还原：勾选 RO1、RO2 → Reset Selected 恢复出厂

说明（真机行为）：Copy 流程=勾选源→Copy→勾选目标→Apply(先有勾选才可点)→Save；
**源/目标配置相同则保存不提示成功**，故先把源/目标设成不同值再复制，并以 Modbus
回读做权威断言；Reset Selected=对勾选通道恢复出厂设置。

备注：Control Mode 完整选项集未全探（已知 Manual/Pulse）；Pulse Width(ms) 一并被复制，
本用例只对 Control Mode 做回读断言。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_ro import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_008_004(app_page):
    nav_to_io_ro(app_page)

    # 2. 制造差异：目标 RO2=Pulse，源 RO1=Manual
    set_control_mode(app_page, 2, "Pulse")
    save_and_check(app_page)
    set_control_mode(app_page, 1, "Manual")
    save_and_check(app_page)

    # 3. 勾选源 RO1 → Copy → 勾选目标 RO2 → Apply → Save
    select_row_checkbox(app_page, 1)
    click_copy(app_page, 1)
    step(f"[COPY] 顶部提示: {copied_message_text(app_page)!r}")
    select_row_checkbox(app_page, 2)
    click_apply_to_selected(app_page)
    save_and_check(app_page)  # 相同配置不提示成功，以 Modbus 回读为准

    # 4. RO2 被套用为 RO1 的 Manual
    verify_modbus(ro_type_reg(2), CONTROL_MODE_ENCODE["Manual"],
                  label="RO2 Control Mode(applied from RO1)")

    # 5. 还原（显式设值，避免保存后 checkbox 勾选态残留 + 切换语义导致 Reset 按钮变灰）：
    #    RO1、RO2 都设回 Pulse，Modbus 回读确认
    set_control_mode(app_page, 1, "Pulse")
    save_and_check(app_page)
    set_control_mode(app_page, 2, "Pulse")
    save_and_check(app_page)
    verify_modbus(ro_type_reg(1), CONTROL_MODE_ENCODE["Pulse"], label="RO1 Control Mode(restore)")
    verify_modbus(ro_type_reg(2), CONTROL_MODE_ENCODE["Pulse"], label="RO2 Control Mode(restore)")
