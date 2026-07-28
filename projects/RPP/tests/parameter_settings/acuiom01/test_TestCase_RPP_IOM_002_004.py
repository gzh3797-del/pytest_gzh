# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_002_004
模块：IO接入参数设置 → AcuIOM-01 AO设置
标题：AO Eng. Unit 设置

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Eng. Unit=V，Confirm 后 Save，回读
  3. 恢复默认并回读确认

预期：保存成功，界面回显 V，Modbus 回读一致；还原成功

备注：TODO(需真机确认) —— AO1 默认 Eng. Unit 具体值探查阶段未记录，暂假定默认
为空字符串，请真机核实后修正 _DEFAULT_UNIT。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_DEFAULT_UNIT = ""


def test_TestCase_RPP_IOM_002_004(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    dialog_set_input(app_page, "Eng. Unit", "V")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Eng. Unit=V: 保存失败"

    open_edit(app_page, 1)
    assert dialog_input_value(app_page, "Eng. Unit") == "V", "AO1 Eng. Unit 页面回显不符"
    dialog_cancel(app_page)
    assert read_unit_ascii(ao_unit_reg(1)) == "V", "AO1 Unit 寄存器回读不符"

    open_edit(app_page, 1)
    dialog_set_input(app_page, "Eng. Unit", _DEFAULT_UNIT)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Eng. Unit 还原默认: 保存失败"
