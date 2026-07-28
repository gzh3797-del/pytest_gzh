# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_001_003
模块：IO接入参数设置 → AcuIOM-01 AI设置
标题：AI Input Lower/Upper Limit 有效边界值

步骤：
  1. 进入 IO-AI 标签，点击 AI1 行 Edit
  2. Input Lower=0.000，Upper=10.000，Confirm 后 Save
  3. 经 Modbus 回读 AI1 Bot Limit(12291)、Top Limit(12289)
  4. 恢复默认 Lower=2.000/Upper=10.000 回读确认

预期：保存成功，回显 Lower=0.000/Upper=10.000；Modbus 回读与设置一致(x0.001)；还原成功

备注：AI1 默认 Lower/Upper=2.000/10.000 为探查阶段真机实测原始值（IOM01P178S06 AI1）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403

_DEFAULT_LOWER = "2.000"
_DEFAULT_UPPER = "10.000"


def test_TestCase_RPP_IOM_001_003(app_page):
    nav_to_io_ai(app_page)
    open_edit(app_page, 1)

    dialog_set_input(app_page, "Input Lower Limit", "0.000")
    dialog_set_input(app_page, "Input Upper Limit", "10.000")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Lower=0.000/Upper=10.000: 保存失败"

    verify_modbus_scaled(ai_bot_limit_reg(1), 0.000, label="AI1 Bot Limit")
    verify_modbus_scaled(ai_top_limit_reg(1), 10.000, label="AI1 Top Limit")

    # 4. 还原默认值
    open_edit(app_page, 1)
    dialog_set_input(app_page, "Input Lower Limit", _DEFAULT_LOWER)
    dialog_set_input(app_page, "Input Upper Limit", _DEFAULT_UPPER)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 还原默认 Lower/Upper: 保存失败"
    verify_modbus_scaled(ai_bot_limit_reg(1), float(_DEFAULT_LOWER), label="AI1 Bot Limit(restore)")
    verify_modbus_scaled(ai_top_limit_reg(1), float(_DEFAULT_UPPER), label="AI1 Top Limit(restore)")
