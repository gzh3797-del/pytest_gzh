# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_001_07_case41
模块：接入设备参数设置 → AcuRev4100
标题：接线方式为 2 Element 3 Wire Network，Voltage Assignment 和 User Channel 验证

预期（用例表）：接线方式 2 Element 3 Wire Network 下 Voltage Assignment 与 User Channel
配置正确（Va/Vb/Vc 循环，input 与 user 对应关系正确）。

自动化覆盖：切 Wiring→2E3W Network，Modbus 回读接线寄存器 + 全 24 通道 VA（Va/Vb/Vc 循环）；
页面侧确认 VA 可选项={Va,Vb,Vc}、User 表 12 行、UC1 Phase 可编辑。
说明：逐 input/user 遍历配置与"与上位机一致"回读（用例表 step2~4）因当前设备通道池已占满、
Phase 下拉候选耗尽，暂不纳入自动化，保留手工。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_MODE = "2 Element 3 Wire Network"


def test_001_07_case41(app_page, restore_wiring):
    nav_to_user_and_ct(app_page)
    switch_wiring(app_page, _MODE)

    # 1. 接线方式已生效
    verify_wiring_reg(_MODE)

    # 2. VA 规律：Va/Vb/Vc 三通道循环（页面抽样 + 全通道寄存器）
    assert read_va_cell_text(app_page, 0) == "Va", "row1 VA 应为 Va"
    assert read_va_cell_text(app_page, 1) == "Vb", "row2 VA 应为 Vb"
    assert read_va_cell_text(app_page, 2) == "Vc", "row3 VA 应为 Vc"
    verify_va_pattern(_MODE)

    # 3. VA 可选项 = {Va, Vb, Vc}
    opts = set(open_va_options(app_page, 0))
    assert opts == {"Va", "Vb", "Vc"}, f"row1 VA 可选项应为 {{Va,Vb,Vc}}，实际 {opts}"

    # 4. User and Channel Mapping 显示且 12 行，UC1 Phase A 可编辑
    assert user_mapping_visible(app_page), "2E3W Network 下 User Channel 区块应显示"
    assert user_channel_row_count(app_page) == 12, "User Channel 应为 12 行"
    assert not phase_is_disabled(app_page, 0, "A"), "UC1 Phase A 应可编辑"
