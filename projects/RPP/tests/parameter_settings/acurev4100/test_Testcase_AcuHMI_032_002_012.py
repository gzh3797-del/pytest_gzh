# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_012
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：修改名称后系统内立即生效

预期：改 UC4 Description=name4 保存后，不重启/不重新登录，打开另一引用该名的页面即刻更新。
验证：User and CT 改 UC4→Save（Modbus 回读确认）→ 同会话切到 Metering→Realtime，
      Meter Point 下拉选 User Channel 4，显示文本立即含 name4（格式 "User Channel 4:name4"）。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_NAME = "name4"


def test_032_002_012(app_page):
    # 1. 改名并保存
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)
    set_description(app_page, 3, _NAME)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"UC4 改名应保存成功，实得警告：{msg!r}"
    verify_description_modbus(4, _NAME)

    # 2. 同会话切到引用页，立即生效（无需刷新/重登）
    nav_to_metering(app_page, "Realtime")
    shown = select_meter_point_uc(app_page, 4)
    assert _NAME in shown, f"Metering→Realtime 应立即显示 {_NAME!r}，实得 {shown!r}"
