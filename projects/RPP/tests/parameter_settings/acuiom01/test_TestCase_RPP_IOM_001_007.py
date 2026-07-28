# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_001_007
模块：IO接入参数设置 → AcuIOM-01 AI设置
标题：AI Breakpoints 分段线性点下发回读

步骤：
  1. 进入 IO-AI 标签，点击 AI1 行 Edit
  2. Segments=3，填 Point1~4 Input(V)=0/3/6/10、Eng.Value=0/30/60/100，Confirm 后 Save
  3. 经 Modbus 回读 AI1 Line Num(12293)、PointX1-4(12294..)、PointY1-4(12302..)
  4. 恢复默认断点并回读确认

预期：保存成功，折线图联动；Modbus 回读 Line Num=3、断点 X/Y 与设置一致(x0.001)；还原成功

备注：AI1 默认 Segments=2、Point1~4=(2,2)/(3,3)/(4,4)/(5,5 禁用) 为探查阶段真机
实测原始值（IOM01P178S06 AI1）。列序：AI 为 col1=Input(V)，col2=Eng.Value(x)。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom01 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403

_NEW_X = [0, 3, 6, 10]
_NEW_Y = [0, 30, 60, 100]
_DEFAULT_SEGMENTS = 2
_DEFAULT_POINTS = [(2.000, 2.000), (3.000, 3.000), (4.000, 4.000), (5.000, 5.000)]


def test_TestCase_RPP_IOM_001_007(app_page):
    nav_to_io_ai(app_page)
    open_edit(app_page, 1)

    set_segments(app_page, 3)
    for idx, point in enumerate((1, 2, 3, 4)):
        set_breakpoint(app_page, point, 1, _NEW_X[idx])
        set_breakpoint(app_page, point, 2, _NEW_Y[idx])
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Breakpoints Segments=3: 保存失败"

    verify_modbus(ai_line_num_reg(1), 3, label="AI1 Line Num")
    for idx, point in enumerate((1, 2, 3, 4)):
        verify_modbus_scaled(ai_point_x_reg(1, point), _NEW_X[idx], label=f"AI1 PointX{point}")
        verify_modbus_scaled(
            ai_point_y_reg(1, point), _NEW_Y[idx], signed=True, label=f"AI1 PointY{point}"
        )

    # 4. 恢复默认断点
    open_edit(app_page, 1)
    set_segments(app_page, _DEFAULT_SEGMENTS)
    for point, (x_val, y_val) in zip((1, 2, 3), _DEFAULT_POINTS[:3]):
        set_breakpoint(app_page, point, 1, x_val)
        set_breakpoint(app_page, point, 2, y_val)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AI1 Breakpoints 还原默认: 保存失败"

    verify_modbus(ai_line_num_reg(1), _DEFAULT_SEGMENTS, label="AI1 Line Num(restore)")
    for point, (x_val, y_val) in zip((1, 2, 3), _DEFAULT_POINTS[:3]):
        verify_modbus_scaled(ai_point_x_reg(1, point), x_val, label=f"AI1 PointX{point}(restore)")
        verify_modbus_scaled(
            ai_point_y_reg(1, point), y_val, signed=True, label=f"AI1 PointY{point}(restore)"
        )
