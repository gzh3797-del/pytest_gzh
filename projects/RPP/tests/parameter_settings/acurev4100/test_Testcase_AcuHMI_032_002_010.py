# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_010
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：名称长度 20 可保存且 21 禁止保存

预期：20 个 ASCII 字符可保存成功；21 个字符保存失败并提示长度须 ≤ 20。
验证：20 字符 → Save 成功 → Modbus 一致；21 字符 → Save 出现 warning 且文案含 "less than 20"，
      Modbus 保持 20 字符那次的值（未被覆盖）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_NAME_20 = "UC123456789012345678"       # 20 字符
_NAME_21 = "UC1234567890123456789"      # 21 字符


def test_032_002_010(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    # 20 字符：可保存
    assert len(_NAME_20) == 20
    set_description(app_page, 0, _NAME_20)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"20 字符应保存成功，实得警告：{msg!r}"
    verify_description_modbus(1, _NAME_20)

    # 21 字符：禁止保存
    assert len(_NAME_21) == 21
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)
    set_description(app_page, 0, _NAME_21)
    warning, msg = save_user_and_ct(app_page)
    assert warning and ERR_DESC_LENGTH.lower() in msg.lower(), \
        f"21 字符应提示长度须 ≤20，实得：warning={warning} msg={msg!r}"
    # 未覆盖：仍为上一步保存的 20 字符值
    verify_description_modbus(1, _NAME_20)
