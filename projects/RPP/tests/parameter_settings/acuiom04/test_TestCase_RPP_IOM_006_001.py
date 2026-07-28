# -*- coding: utf-8 -*-
"""
TestCase_RPP_IOM_006_001
模块：IO接入参数设置 → AcuIOM-04 DI设置
标题：DI 页面进入与 Function 保存冒烟

步骤：
  1. 进入 Settings-Devices-AcuIOM-04-点顶部 IO-DI 标签
  2. DI1 Function 选 Status Monitor，Save
  3. 界面回读 DI1 Function

预期：
  1. 成功进入 AcuIOM-04 IO-DI 标签，展示 28 行(DI 1~28)
  2. 保存成功
  3. 界面回显 Status Monitor
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from helpers_iom04 import *  # noqa: F401, F403
from _src_io_di import *  # noqa: F401, F403


def test_TestCase_RPP_IOM_006_001(app_page):
    nav_to_io_di(app_page)

    assert di_row_count(app_page) == DI_COUNT, f"DI 标签应展示 {DI_COUNT} 行"

    set_function(app_page, 1, "Status Monitor")
    saved = save_and_check(app_page)
    assert saved, "DI1 Function=Status Monitor: 保存失败"

    assert get_function_text(app_page, 1) == "Status Monitor", "DI1 Function 页面回显不符"
