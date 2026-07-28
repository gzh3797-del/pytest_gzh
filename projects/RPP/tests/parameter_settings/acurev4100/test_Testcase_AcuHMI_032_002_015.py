# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_015
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：批量修改多个通道后显示与保存一致

预期：UC1/UC2/UC12 三通道批量改名保存成功；刷新页面、重新登录后显示仍保持。
验证：批量填 → Save → 重新导航（刷新页面）确认输入框回显一致 + Modbus 回读一致。
说明：设备侧持久化由 Modbus 回读权威证明（重登后设备寄存器值不变即等价）；用例表 step3
"退出重登"因 app_page 为会话级共享登录、单独登出会干扰其它用例，此处以"重导航刷新 + Modbus
持久化"覆盖其意图，独立的登出重登验证保留手工。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_BATCH = {0: "name1", 1: "name2", 11: "name12"}   # uc_idx → name（UC1/UC2/UC12）


def test_032_002_015(app_page):
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)

    for uc_idx, name in _BATCH.items():
        set_description(app_page, uc_idx, name)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"批量改名应保存成功，实得警告：{msg!r}"

    # 刷新页面（重新导航）后回显仍一致
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)
    for uc_idx, name in _BATCH.items():
        shown = read_description_input(app_page, uc_idx)
        assert shown == name, f"UC{uc_idx + 1} 刷新后回显应为 {name!r}，实得 {shown!r}"

    # Modbus 回读确认设备侧持久化（等价重登后仍保持）
    for uc_idx, name in _BATCH.items():
        verify_description_modbus(uc_idx + 1, name)
