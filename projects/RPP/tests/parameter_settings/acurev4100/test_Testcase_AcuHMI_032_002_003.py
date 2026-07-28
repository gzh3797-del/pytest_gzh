# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_003
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：12 个 User Channel 可独立配置名称

预期：UC1=Name1、UC2=Name2 可独立保存，互不覆盖，其他通道不受影响。
验证：填 UC1/UC2 → Save → Modbus 回读各自 Description 寄存器一致且互不串。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403


def test_032_002_003(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    set_description(app_page, 0, "Name1")
    set_description(app_page, 1, "Name2")
    save_user_and_ct(app_page)

    # 独立保存、互不覆盖（Modbus 回读为准）
    verify_description_modbus(1, "Name1")
    verify_description_modbus(2, "Name2")
