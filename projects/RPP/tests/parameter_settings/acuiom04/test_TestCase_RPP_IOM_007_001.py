# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_007_001
模块：IO接入参数设置 → AcuIOM-04 DO设置
标题：DO Control Mode 切换

步骤：
  1. 进入 IO-DO 标签
  2. DO1 Control Mode 依次设为 Pulse、Manual，每次 Save，界面回读
  3. 恢复默认并回读确认

预期：各模式保存成功，界面回显与所选一致；还原成功

备注：Control Mode 完整选项集未在探查中抓全(仅见 Pulse/Manual)，实测以页面下拉
为准补充。DO1 默认 Control Mode="Pulse" 为探查阶段真机实测原始值（IOM04P193S02）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_do import *  # noqa: F401, F403

_MODES = ["Pulse", "Manual"]
_DEFAULT_MODE = "Pulse"


def test_TestCase_RPP_IOM_007_001(app_page):
    nav_to_io_do(app_page)
    assert do_row_count(app_page) == DO_COUNT, f"DO 标签应展示 {DO_COUNT} 行"

    for option in _MODES:
        set_control_mode(app_page, 1, option)
        saved = save_and_check(app_page)
        assert saved, f"DO1 Control Mode={option}: 保存失败"
        assert get_control_mode_text(app_page, 1) == option, f"DO1 Control Mode 回显应为 {option}"

    set_control_mode(app_page, 1, _DEFAULT_MODE)
    saved = save_and_check(app_page)
    assert saved, "DO1 还原默认 Control Mode: 保存失败"
    assert get_control_mode_text(app_page, 1) == _DEFAULT_MODE, "DO1 还原后回显不符"
