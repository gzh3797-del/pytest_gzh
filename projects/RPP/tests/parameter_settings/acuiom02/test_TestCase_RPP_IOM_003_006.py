# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_003_006
模块：IO接入参数设置 → AcuIOM-02 AI设置
标题：AI Number of Segments 联动断点表可编辑行

步骤：
  1. 进入 IO-AI 标签，点击 AI1 行 Edit
  2. Segments 选 2 观察断点行；选 3 再观察
  3. Cancel 关闭(不保存)

预期：Segments=2 时 Point4 禁用；Segments=3 时 Point1~4 全可编辑；弹窗关闭未改动设备

本用例全程不点 Save，仅弹窗内 Cancel 关闭——不改动设备任何配置。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ai import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_003_006(app_page):
    nav_to_io_ai(app_page)
    open_edit(app_page, 1)

    set_segments(app_page, 2)
    for point in (1, 2, 3):
        assert not breakpoint_disabled(app_page, point, 1), f"Segments=2 时 Point{point} X 应可编辑"
        assert not breakpoint_disabled(app_page, point, 2), f"Segments=2 时 Point{point} Y 应可编辑"
    assert breakpoint_disabled(app_page, 4, 1), "Segments=2 时 Point4 X 应禁用"
    assert breakpoint_disabled(app_page, 4, 2), "Segments=2 时 Point4 Y 应禁用"

    set_segments(app_page, 3)
    for point in (1, 2, 3, 4):
        assert not breakpoint_disabled(app_page, point, 1), f"Segments=3 时 Point{point} X 应可编辑"
        assert not breakpoint_disabled(app_page, point, 2), f"Segments=3 时 Point{point} Y 应可编辑"

    dialog_cancel(app_page)
