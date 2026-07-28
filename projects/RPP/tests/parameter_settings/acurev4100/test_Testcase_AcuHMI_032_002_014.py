# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_014
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：未修改通道保持默认名称

预期：仅改 UC1，UC2~12 仍保持原名不受影响。
验证：先记录 UC2/UC12 现值 → 仅改 UC1 → Save → Modbus 确认 UC1 更新、UC2/UC12 未变。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_NAME = "name1"


def test_032_002_014(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    # 记录未涉及通道的当前值（代表：UC2、UC12）
    before_uc2 = read_description_modbus(2)
    before_uc12 = read_description_modbus(12)

    set_description(app_page, 0, _NAME)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"UC1 应保存成功，实得警告：{msg!r}"

    verify_description_modbus(1, _NAME)
    # 未修改通道保持原值
    verify_description_modbus(2, before_uc2)
    verify_description_modbus(12, before_uc12)
