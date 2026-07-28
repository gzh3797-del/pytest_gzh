# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_008
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：ASCII 合法名称可保存

预期：UC1 Description = "MainFeed_A1-01" 保存成功，显示与输入一致。
验证：填入 → Save → Modbus 回读一致。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_NAME = "MainFeed_A1-01"


def test_032_002_008(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    set_description(app_page, 0, _NAME)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"合法 ASCII 名称应保存成功，实得警告：{msg!r}"

    verify_description_modbus(1, _NAME)
