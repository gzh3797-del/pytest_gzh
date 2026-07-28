# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_001_005
模块：IO接入参数设置 → AcuIOM-01 AI设置
标题：AI Eng. Unit 设置

步骤：
  1. 进入 IO-AI 标签，点击 AI1 行 Edit
  2. Eng. Unit=kW，Confirm 后 Save，回读
  3. 恢复默认并回读确认

预期：保存成功；界面回显 kW；Modbus 回读单位区(12800)一致；还原成功

备注：AI1 默认 Eng. Unit="x" 为探查阶段真机实测原始值（IOM01P178S06 AI1）。
TODO(需真机确认)：Unit 寄存器(uint32/2 regs)按 ASCII 打包推测解析（见
helpers_iom01.read_unit_ascii 注释），未做真实回读比对验证。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403

_DEFAULT_UNIT = "x"


def test_TestCase_RPP_IOM_001_005(app_page):
    nav_to_io_ai(app_page)
    open_edit(app_page, 1)

    dialog_set_input(app_page, "Eng. Unit", "kW")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Eng. Unit=kW: 保存失败"

    open_edit(app_page, 1)
    assert dialog_input_value(app_page, "Eng. Unit") == "kW", "AI1 Eng. Unit 页面回显不符"
    dialog_cancel(app_page)
    assert read_unit_ascii(ai_unit_reg(1)) == "kW", "AI1 Unit 寄存器回读不符"

    # 3. 恢复默认
    open_edit(app_page, 1)
    dialog_set_input(app_page, "Eng. Unit", _DEFAULT_UNIT)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Eng. Unit 还原默认: 保存失败"
    assert read_unit_ascii(ai_unit_reg(1)) == _DEFAULT_UNIT, "AI1 Unit 寄存器还原后不符"
