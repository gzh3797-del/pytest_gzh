# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_004_002
模块：IO接入参数设置 → AcuIOM-02 AO设置
标题：AO Output Lower/Upper Limit 有效边界值

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Output Lower=0.000、Upper=10.000，Confirm 后 Save，Modbus 回读
  3. 恢复默认并回读确认

预期：保存成功，回显与 Modbus 回读一致；还原成功

备注：AcuIOM-02 AO1 出厂 Lower/Upper 具体值探查阶段未记录，本用例改用"运行期先
捕获当前值、用后再还原"策略（而非硬编码假定默认值）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_004_002(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    original_lower = dialog_input_value(app_page, "Output Lower Limit")
    original_upper = dialog_input_value(app_page, "Output Upper Limit")

    dialog_set_input(app_page, "Output Lower Limit", "0.000")
    dialog_set_input(app_page, "Output Upper Limit", "10.000")
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Lower=0.000/Upper=10.000: 保存失败"

    verify_modbus_scaled(ao_bot_limit_reg(1), 0.000, label="AO1 Bot Limit")
    verify_modbus_scaled(ao_top_limit_reg(1), 10.000, label="AO1 Top Limit")

    # 3. 用运行期捕获的原始值还原
    open_edit(app_page, 1)
    dialog_set_input(app_page, "Output Lower Limit", original_lower)
    dialog_set_input(app_page, "Output Upper Limit", original_upper)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 还原原始 Lower/Upper: 保存失败"
    verify_modbus_scaled(ao_bot_limit_reg(1), float(original_lower), label="AO1 Bot Limit(restore)")
    verify_modbus_scaled(ao_top_limit_reg(1), float(original_upper), label="AO1 Top Limit(restore)")
