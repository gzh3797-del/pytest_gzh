# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_006
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：Description 空值回退显示规则正确

预期：允许空值保存（不强制命名）；Description 为空时界面显示回退为默认名 "User Channel N"。
验证：清空 UC3 Description → Save（成功）→ Modbus 回读为空 + 页面回退显示 "User Channel 3"。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403


def test_032_002_006(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    set_description(app_page, 2, "")   # UC3 置空
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"空值应允许保存，实得警告：{msg!r}"

    # 1. 允许空值：Modbus 回读为空
    verify_description_modbus(3, "")
    # 2. 引用界面回退显示默认名 "User Channel 3"（空值不在 Description 输入框回退，
    #    而是在展示页如 Metering→Realtime 的 Meter Point 处回退）
    nav_to_metering(app_page, "Realtime")
    shown = select_meter_point_uc(app_page, 3)
    assert "User Channel 3" in shown, f"空值回退显示应含 'User Channel 3'，实得 {shown!r}"
