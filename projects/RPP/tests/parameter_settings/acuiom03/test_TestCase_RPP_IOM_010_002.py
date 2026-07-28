# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_010_002
模块：IO接入参数设置 → AcuIOM-03 RO设置
标题：RO Pulse Width(ms) 有效值下发回读

步骤：
  1. 进入 IO-RO 标签，RO1 Control Mode=Pulse
  2. Pulse Width=2000 ms，Save
  3. 界面回读；Modbus 回读 RO1 Pulse Width 配置寄存器
  4. 恢复默认值回读确认

预期：保存成功；界面回显 2000，Modbus 回读一致；还原成功

备注：AcuIOM-03 离线，需上线后执行（conftest 自动 pytest.skip）。RO1 默认 Pulse Width
待上线核实（暂按 IOM-04 同款 2000ms）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom03 import *  # noqa: F401, F403
from _src_io_ro import *  # noqa: F401, F403

_DEFAULT_WIDTH = 2000


def test_TestCase_RPP_IOM_010_002(app_page):
    nav_to_io_ro(app_page)
    set_control_mode(app_page, 1, "Pulse")

    set_pulse_width(app_page, 1, "2000")
    saved = save_and_check(app_page)
    assert saved, "RO1 Pulse Width=2000: 保存失败"

    assert pulse_width_value(app_page, 1) == "2000", "RO1 Pulse Width 页面回显不符"
    verify_modbus(ro_pulse_width_reg(1), 2000, label="RO1 Pulse Width")

    set_pulse_width(app_page, 1, str(_DEFAULT_WIDTH))
    saved = save_and_check(app_page)
    assert saved, "RO1 Pulse Width 还原默认: 保存失败"
    verify_modbus(ro_pulse_width_reg(1), _DEFAULT_WIDTH, label="RO1 Pulse Width(restore)")
