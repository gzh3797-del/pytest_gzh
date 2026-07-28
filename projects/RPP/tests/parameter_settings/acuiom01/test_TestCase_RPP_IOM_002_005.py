# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_002_005
模块：IO接入参数设置 → AcuIOM-01 AO设置
标题：AO Breakpoints 分段线性点(Eng.Value→Output)下发回读

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Segments=3，按列序 Eng.Value(x)->Output(V) 填 Point1~4，Confirm 后 Save
  3. 经 Modbus 回读 AO1 Line Num、断点 X/Y
  4. 恢复默认断点并回读确认

预期：保存成功，折线图联动(注意 AO 列序与 AI 相反)；Modbus 回读一致；还原成功

备注：AO1 默认 Number of Segments=1 为探查阶段真机实测原始值（IOM01P178S06 AO1，
`.el-select` 回显文本 "1"）。TODO(需真机确认) —— 默认断点 Point 数值探查阶段未
记录（弹窗 innerText 转储不含 input value），本用例第 4 步仅还原 Segments=1，
断点数值保持本用例第 2 步下发值不做进一步还原，请真机核实 AO1 出厂断点值后补全。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_NEW_X = [0, 30, 60, 100]   # Eng. Value(x) 列
_NEW_Y = [0, 3, 6, 10]      # Output(V) 列
_DEFAULT_SEGMENTS = 1


def test_TestCase_RPP_IOM_002_005(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    set_segments(app_page, 3)
    for idx, point in enumerate((1, 2, 3, 4)):
        set_breakpoint(app_page, point, 1, _NEW_X[idx])
        set_breakpoint(app_page, point, 2, _NEW_Y[idx])
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Breakpoints Segments=3: 保存失败"

    verify_modbus(ao_line_num_reg(1), 3, label="AO1 Line Num")
    for idx, point in enumerate((1, 2, 3, 4)):
        verify_modbus_scaled(
            ao_point_x_reg(1, point), _NEW_X[idx], signed=True, label=f"AO1 PointX{point}"
        )
        verify_modbus_scaled(ao_point_y_reg(1, point), _NEW_Y[idx], label=f"AO1 PointY{point}")

    # 4. 恢复默认 Segments（断点数值保持本用例下发值，出厂值待真机确认后补全还原）
    open_edit(app_page, 1)
    set_segments(app_page, _DEFAULT_SEGMENTS)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Segments 还原默认: 保存失败"
    verify_modbus(ao_line_num_reg(1), _DEFAULT_SEGMENTS, label="AO1 Line Num(restore)")
