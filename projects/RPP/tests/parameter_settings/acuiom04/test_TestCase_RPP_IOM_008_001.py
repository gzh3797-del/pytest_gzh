# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_008_001
模块：IO接入参数设置 → AcuIOM-04 RO设置
标题：RO Control Mode 切换

步骤：
  1. 进入 IO-RO 标签
  2. RO1 Control Mode 依次设为 Pulse、Manual，每次 Save，界面回读
  3. 恢复默认并回读确认

预期：各模式保存成功，界面回显与所选一致；还原成功

备注：Control Mode 完整选项集未在探查中抓全(仅见 Pulse/Manual)，实测以页面下拉
为准补充。RO1 默认 Control Mode="Pulse" 为探查阶段真机实测原始值（IOM04P193S02）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_ro import *  # noqa: F401, F403

_MODES = ["Pulse", "Manual"]
_DEFAULT_MODE = "Pulse"


def test_TestCase_RPP_IOM_008_001(app_page):
    nav_to_io_ro(app_page)
    assert ro_row_count(app_page) == RO_COUNT, f"RO 标签应展示 {RO_COUNT} 行"

    for option in _MODES:
        set_control_mode(app_page, 1, option)
        saved = save_and_check(app_page)
        assert saved, f"RO1 Control Mode={option}: 保存失败"
        assert get_control_mode_text(app_page, 1) == option, f"RO1 Control Mode 回显应为 {option}"

    set_control_mode(app_page, 1, _DEFAULT_MODE)
    saved = save_and_check(app_page)
    assert saved, "RO1 还原默认 Control Mode: 保存失败"
    assert get_control_mode_text(app_page, 1) == _DEFAULT_MODE, "RO1 还原后回显不符"
