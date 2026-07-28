# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_004_003
模块：IO接入参数设置 → AcuIOM-02 AO设置
标题：AO Output Limit 非法值校验

步骤：
  1. 进入 IO-AO 标签，点击 AO1 行 Edit
  2. Output Upper Limit 依次输入：空、abc、@、-1、超上界，逐一 Confirm/Save

预期：每种非法输入均被阻止保存并提示，配置不写入设备

备注：TODO(需真机确认) —— 仅断言"存在可见校验错误或弹窗未关闭"，超上界样例值为占位。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom02 import *  # noqa: F401, F403
from _src_io_ao import *  # noqa: F401, F403

_INVALID_VALUES = ["", "abc", "@", "-1", "99999.999"]


def test_TestCase_RPP_IOM_004_003(app_page):
    nav_to_io_ao(app_page)
    open_edit(app_page, 1)

    for bad_value in _INVALID_VALUES:
        dialog_set_input(app_page, "Output Upper Limit", bad_value)
        dialog_confirm(app_page)
        app_page.wait_for_timeout(500)
        errors = get_visible_errors(app_page)
        dialog_still_open = _dialog(app_page).count() > 0 and _dialog(app_page).first.is_visible()
        assert errors or dialog_still_open, (
            f"Output Upper Limit={bad_value!r}: 期望被阻止保存，但未检测到任何阻止迹象"
        )

    if _dialog(app_page).count() > 0 and _dialog(app_page).first.is_visible():
        dialog_cancel(app_page)
