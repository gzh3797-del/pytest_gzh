# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_002_001
模块：IO接入参数设置 → AcuIOM-01 AO设置
标题：AO Signal Type 四种输出类型遍历

步骤：
  1. 进入 IO-AO 标签
  2. AO1 Signal Type 依次设为四种类型，每次 Save
  3. 每次经 Modbus 回读 AO1 Type 寄存器
  4. 恢复 AO1=Voltage(0-10V) 回读确认

预期：四种类型均保存成功；Modbus 回读编码值与所选一致；还原成功
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_TYPES = ["Voltage(0-10V)", "Voltage(2-10V)", "Current(0-20mA)", "Current(4-20mA)"]


def test_TestCase_RPP_IOM_002_001(app_page):
    nav_to_io_ao(app_page)
    assert ao_row_count(app_page) == AO_COUNT, f"AO 标签应展示 {AO_COUNT} 行"

    for option in _TYPES:
        set_signal_type(app_page, 1, option)
        saved = save_and_check(app_page)
        assert saved, f"AO1 Signal Type={option}: 保存失败"
        verify_modbus(ao_type_reg(1), SIGNAL_TYPE_ENCODE[option], label=f"AO1 Type({option})")

    set_signal_type(app_page, 1, "Voltage(0-10V)")
    saved = save_and_check(app_page)
    assert saved, "AO1 还原 Voltage(0-10V): 保存失败"
    verify_modbus(ao_type_reg(1), SIGNAL_TYPE_ENCODE["Voltage(0-10V)"], label="AO1 Type(restore)")
