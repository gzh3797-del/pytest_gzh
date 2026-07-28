# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_009
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：非 ASCII 字符禁止保存

预期：Description 含中文（"主回路1"、"Name中1"）保存失败，提示仅允许 ASCII 字符。
验证：填非 ASCII → Save → 出现 .el-message--warning 且文案含 "must contain only ASCII"；
      Modbus 回读确认未被写入（保持原值）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403


def _expect_reject(app_page, text):
    before = read_description_modbus(1)
    set_description(app_page, 0, text)
    warning, msg = save_user_and_ct(app_page)
    assert warning and ERR_DESC_ASCII.lower() in msg.lower(), \
        f"非 ASCII {text!r} 应提示仅允许 ASCII，实得：warning={warning} msg={msg!r}"
    # 未持久化：Modbus 保持原值
    after = read_description_modbus(1)
    assert after == before, f"非法值不应写入，before={before!r} after={after!r}"


def test_032_002_009(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    _expect_reject(app_page, "主回路1")
    nav_to_user_and_ct(app_page)          # 重载页面清掉未保存的非法输入
    ensure_user_mapping(app_page)
    _expect_reject(app_page, "Name中1")
