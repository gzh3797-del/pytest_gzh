# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_005_006
模块：IO接入参数设置 → AcuIOM-03 DI设置
标题：DI Copy 批量复制 + Reset 恢复出厂

步骤：
  1. 进入 IO-DI 标签
  2. 前置：源 DI1 Function 设为 Pulse Counter（区别于默认 Status Monitor），保存
  3. 勾选源 DI1 → Copy → 勾选目标 DI2 → Apply to Selected → Save
  4. Modbus 回读 DI2 Function，确认被套用为 DI1 的 Pulse Counter
  5. 还原：DI2 Reset 出厂；DI1 设回 Status Monitor

说明（真机行为）：Copy 流程=勾选源→Copy→勾选目标→Apply(先有勾选才可点)→Save；
**源/目标配置相同则保存不提示成功**，故先制造差异并以 Modbus 回读做权威断言；
Reset Selected=对勾选通道恢复出厂设置。

备注：AcuIOM-03(IOM03P170S04)当前离线，本用例需设备上线后执行（conftest 会 pytest.skip）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_005_006(app_page):
    nav_to_io_di(app_page)

    set_function(app_page, 1, "Pulse Counter")
    save_and_check(app_page)

    select_row_checkbox(app_page, 1)
    click_copy(app_page, 1)
    step(f"[COPY] 顶部提示: {copied_message_text(app_page)!r}")
    select_row_checkbox(app_page, 2)
    click_apply_to_selected(app_page)
    save_and_check(app_page)  # 相同配置不提示成功，以 Modbus 回读为准

    verify_modbus(di_type_reg(2), FUNCTION_ENCODE["Pulse Counter"],
                  label="DI2 Function(applied from DI1)")

    select_row_checkbox(app_page, 2)
    click_reset_selected(app_page)
    save_and_check(app_page)
    set_function(app_page, 1, "Status Monitor")
    save_and_check(app_page)
    verify_modbus(di_type_reg(1), FUNCTION_ENCODE["Status Monitor"], label="DI1 Function(restore)")
