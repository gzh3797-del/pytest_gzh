# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_001_07_case38
模块：接入设备参数设置 → AcuRev4100
标题：接线方式为 1E2W，Voltage Assignment 和 User Channel 验证

预期（用例表）：接线方式 1 Element 2 Wire 下，input1-24 Voltage Assignment 仅支持 Va，
User Channel 不展示。

自动化覆盖：切 Wiring→1E2W（自动保存），Modbus 回读接线寄存器 + 全 24 通道 VA=Va；
页面侧确认 VA 列显示 Va、VA 下拉只读、User and Channel Mapping 区块不显示。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_MODE = "1 Element 2 Wire"


def test_001_07_case38(app_page, restore_wiring):
    nav_to_user_and_ct(app_page)
    switch_wiring(app_page, _MODE)

    # 1. 接线方式已生效（Modbus）
    verify_wiring_reg(_MODE)

    # 2. Voltage Assignment 全部为 Va —— 页面抽样 + 全 24 通道寄存器
    for r in range(3):
        assert read_va_cell_text(app_page, r) == "Va", f"row{r + 1} VA 显示应为 Va"
    verify_va_pattern(_MODE)

    # 3. VA 下拉只读（本模式仅 Va，不可改）
    assert va_is_disabled(app_page, 0), "1E2W 下 Voltage Assignment 下拉应为 disabled 只读"

    # 4. User and Channel Mapping 区块不显示
    assert not user_mapping_visible(app_page), "1E2W 下 User Channel 区块应不显示"
