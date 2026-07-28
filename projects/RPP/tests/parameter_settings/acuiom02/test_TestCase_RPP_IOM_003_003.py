# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_003_003
模块：IO接入参数设置 → AcuIOM-02 AI设置
标题：AI Input Lower/Upper Limit 有效边界值

步骤：
  1. 进入 IO-AI 标签，点击 AI1 行 Edit
  2. Input Lower=0.000、Upper=10.000，Confirm 后 Save
  3. 经 Modbus 回读 AI1 Bot/Top Limit
  4. 恢复默认并回读确认

预期：保存成功，回显一致；Modbus 回读一致(x0.001)；还原成功

备注：AI1 默认 Lower=0.000/Upper=10.000 为探查阶段真机实测原始值（IOM02P118S01
AI1，与本用例下发值恰好相同，仍完整走一遍下发/回读/还原流程）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403

_DEFAULT_LOWER = "0.000"
_DEFAULT_UPPER = "10.000"


def test_TestCase_RPP_IOM_003_003(app_page):
    nav_to_io_ai(app_page)
    open_edit(app_page, 1)

    dialog_set_input(app_page, "Input Lower Limit", "0.000")
    dialog_set_input(app_page, "Input Upper Limit", "10.000")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Lower=0.000/Upper=10.000: 保存失败"

    verify_modbus_scaled(ai_bot_limit_reg(1), 0.000, label="AI1 Bot Limit")
    verify_modbus_scaled(ai_top_limit_reg(1), 10.000, label="AI1 Top Limit")

    open_edit(app_page, 1)
    dialog_set_input(app_page, "Input Lower Limit", _DEFAULT_LOWER)
    dialog_set_input(app_page, "Input Upper Limit", _DEFAULT_UPPER)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 还原默认 Lower/Upper: 保存失败"
    verify_modbus_scaled(ai_bot_limit_reg(1), float(_DEFAULT_LOWER), label="AI1 Bot Limit(restore)")
    verify_modbus_scaled(ai_top_limit_reg(1), float(_DEFAULT_UPPER), label="AI1 Top Limit(restore)")
