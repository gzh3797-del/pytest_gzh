# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_009_004
模块：IO接入参数设置 → AcuIOM-03 DO设置
标题：DO Copy 批量复制 + 还原

步骤：
  1. 进入 IO-DO 标签
  2. 前置：目标 DO2 Control Mode 设为 Pulse、源 DO1 设为 Manual（制造差异），分别 Save
  3. 勾选源 DO1 → Copy → 勾选目标 DO2 → Apply to Selected → Save
  4. Modbus 回读 DO2 Control Mode，确认被套用为 DO1 的 Manual
  5. 还原：DO1、DO2 都设回 Pulse，Modbus 回读确认

预期：Apply 后 DO2 被套用为 DO1 配置；Modbus 回读 DO2=Manual(0)；还原后 DO1/DO2=Pulse(1)

备注：AcuIOM-03 离线，需上线后执行（conftest 自动 pytest.skip）。Copy 流程=勾选源→Copy→
勾选目标→Apply(先有勾选才可点)→Save；还原用显式设值（保存后复选框勾选态残留 + 切换语义
会使 Reset Selected 变灰，故不用 Reset）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_do import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_009_004(app_page):
    nav_to_io_do(app_page)

    # 2. 制造差异：目标 DO2=Pulse，源 DO1=Manual
    set_control_mode(app_page, 2, "Pulse")
    save_and_check(app_page)
    set_control_mode(app_page, 1, "Manual")
    save_and_check(app_page)

    # 3. 勾选源 DO1 → Copy → 勾选目标 DO2 → Apply → Save
    select_row_checkbox(app_page, 1)
    click_copy(app_page, 1)
    step(f"[COPY] 顶部提示: {copied_message_text(app_page)!r}")
    select_row_checkbox(app_page, 2)
    click_apply_to_selected(app_page)
    save_and_check(app_page)  # 相同配置不提示成功，以 Modbus 回读为准

    # 4. DO2 被套用为 DO1 的 Manual
    verify_modbus(do_type_reg(2), CONTROL_MODE_ENCODE["Manual"],
                  label="DO2 Control Mode(applied from DO1)")

    # 5. 还原：DO1、DO2 都设回 Pulse
    set_control_mode(app_page, 1, "Pulse")
    save_and_check(app_page)
    set_control_mode(app_page, 2, "Pulse")
    save_and_check(app_page)
    verify_modbus(do_type_reg(1), CONTROL_MODE_ENCODE["Pulse"], label="DO1 Control Mode(restore)")
    verify_modbus(do_type_reg(2), CONTROL_MODE_ENCODE["Pulse"], label="DO2 Control Mode(restore)")
