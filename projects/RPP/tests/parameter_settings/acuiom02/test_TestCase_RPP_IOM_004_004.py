# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_004_004
模块：IO接入参数设置 → AcuIOM-02 AO设置
标题：AO Breakpoints 分段线性点下发回读

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Segments=3，按 Eng.Value->Output 列序填 Point1~4，Confirm 后 Save，Modbus 回读
  3. 恢复默认并回读确认

预期：保存成功，折线图联动；Modbus 回读一致；还原成功

备注：AcuIOM-02 AO1 出厂 Segments/断点具体值探查阶段未记录，本用例改用"运行期先
捕获当前值、用后再还原"策略（而非硬编码假定默认值）。AO 列序与 AI 相反：
col1="Eng. Value (x)"，col2="Output (V)"。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_NEW_X = [0, 30, 60, 100]   # Eng. Value(x) 列
_NEW_Y = [0, 3, 6, 10]      # Output(V) 列


def test_TestCase_RPP_IOM_004_004(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    original_segments = dialog_dropdown_text(app_page, "Number of Segments")
    original_points = [
        (breakpoint_value(app_page, p, 1), breakpoint_value(app_page, p, 2))
        for p in (1, 2, 3)
        if not breakpoint_disabled(app_page, p, 1)
    ]
    dialog_cancel(app_page)

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

    # 3. 用运行期捕获的原始值还原
    open_edit(app_page, 1)
    set_segments(app_page, int(original_segments))
    for point, (x_val, y_val) in zip(range(1, len(original_points) + 1), original_points):
        set_breakpoint(app_page, point, 1, x_val)
        set_breakpoint(app_page, point, 2, y_val)
    dialog_confirm(app_page)
    saved = save_and_check(app_page)
    assert saved, "AO1 Breakpoints 还原: 保存失败"
    verify_modbus(ao_line_num_reg(1), int(original_segments), label="AO1 Line Num(restore)")
