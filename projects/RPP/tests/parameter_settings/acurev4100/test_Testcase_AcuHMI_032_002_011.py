# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_011
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：重名与特殊 ASCII 符号处理正确

预期：UC1=Channel_A、UC2=Channel_A（重名）、UC3=U-C_3#A（特殊 ASCII 符号）均可保存，
      展示与输入一致，不乱码/截断/替换。
验证：填 3 个 → Save → Modbus 逐一回读一致。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403


def test_032_002_011(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    set_description(app_page, 0, "Channel_A")
    set_description(app_page, 1, "Channel_A")   # 重名
    set_description(app_page, 2, "U-C_3#A")     # 特殊 ASCII 符号
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"重名/特殊 ASCII 符号应可保存，实得警告：{msg!r}"

    verify_description_modbus(1, "Channel_A")
    verify_description_modbus(2, "Channel_A")
    verify_description_modbus(3, "U-C_3#A")
