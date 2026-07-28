# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_002_002
模块：IO接入参数设置 → AcuIOM-01 AO设置
标题：AO Output Lower/Upper Limit 有效边界值

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Output Lower=0.000、Upper=10.000，Confirm 后 Save
  3. 经 Modbus 回读 AO1 Output Lower/Upper Limit
  4. 恢复默认并回读确认

预期：保存成功，回显 Lower=0.000/Upper=10.000；Modbus 回读一致；还原成功

备注：TODO(需真机确认) —— AO1 默认 Lower/Upper 具体值探查阶段未记录数值（弹窗
innerText 转储不含 input value），本用例假定默认 Lower=0.000/Upper=10.000
（对应 Voltage(0-10V) 常见量程），请真机核实后修正 _DEFAULT_LOWER/_DEFAULT_UPPER。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_DEFAULT_LOWER = "0.000"
_DEFAULT_UPPER = "10.000"


def test_TestCase_RPP_IOM_002_002(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    dialog_set_input(app_page, "Output Lower Limit", "0.000")
    dialog_set_input(app_page, "Output Upper Limit", "10.000")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Lower=0.000/Upper=10.000: 保存失败"

    verify_modbus_scaled(ao_bot_limit_reg(1), 0.000, label="AO1 Bot Limit")
    verify_modbus_scaled(ao_top_limit_reg(1), 10.000, label="AO1 Top Limit")

    open_edit(app_page, 1)
    dialog_set_input(app_page, "Output Lower Limit", _DEFAULT_LOWER)
    dialog_set_input(app_page, "Output Upper Limit", _DEFAULT_UPPER)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 还原默认 Lower/Upper: 保存失败"
    verify_modbus_scaled(ao_bot_limit_reg(1), float(_DEFAULT_LOWER), label="AO1 Bot Limit(restore)")
    verify_modbus_scaled(ao_top_limit_reg(1), float(_DEFAULT_UPPER), label="AO1 Top Limit(restore)")
