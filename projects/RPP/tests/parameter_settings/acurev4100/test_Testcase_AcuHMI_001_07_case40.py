# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_001_07_case40
模块：接入设备参数设置 → AcuRev4100
标题：接线方式为 2 Element 3 Wire Delta，Voltage Assignment 和 User Channel 验证

预期（用例表）：input 奇数行仅 Vab、偶数行仅 Vbc；User 1-12 Phase A/C 有默认 input 且不可选。

自动化覆盖：切 Wiring→2E3W Delta，Modbus 回读接线寄存器 + 全 24 通道 VA（奇 Vab/偶 Vbc）；
页面侧确认 VA 列显示、VA 下拉只读、User 表显示、UC1 Phase A/B/C 均固定(disabled)。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_MODE = "2 Element 3 Wire Delta"


def test_001_07_case40(app_page, restore_wiring):
    nav_to_user_and_ct(app_page)
    switch_wiring(app_page, _MODE)

    # 1. 接线方式已生效
    verify_wiring_reg(_MODE)

    # 2. VA 规律：奇数通道 Vab、偶数通道 Vbc（页面抽样 + 全通道寄存器）
    assert read_va_cell_text(app_page, 0) == "Vab", "row1 VA 应为 Vab"
    assert read_va_cell_text(app_page, 1) == "Vbc", "row2 VA 应为 Vbc"
    verify_va_pattern(_MODE)

    # 3. VA 下拉只读（Delta 下 VA 固定不可改）
    assert va_is_disabled(app_page, 0), "row1 VA 下拉应为 disabled"
    assert va_is_disabled(app_page, 1), "row2 VA 下拉应为 disabled"

    # 4. User and Channel Mapping 显示；Phase A/B/C 固定(默认 input 不可选)
    assert user_mapping_visible(app_page), "2E3W Delta 下 User Channel 区块应显示"
    assert phase_is_disabled(app_page, 0, "A"), "UC1 Phase A 应固定/disabled"
    assert phase_is_disabled(app_page, 0, "B"), "UC1 Phase B 应固定/disabled"
    assert phase_is_disabled(app_page, 0, "C"), "UC1 Phase C 应固定/disabled"
