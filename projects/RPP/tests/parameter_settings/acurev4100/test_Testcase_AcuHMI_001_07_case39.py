# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_001_07_case39
模块：接入设备参数设置 → AcuRev4100
标题：接线方式为 2 Element 3 Wire 1 Phase，Voltage Assignment 和 User Channel 验证

预期（用例表 step2）：input1-24 Voltage Assignment 可选 Va、Vc，User1-12 显示且
Phase A/C 可配置 Input Channel。（step3 负向"非法配置保存失败"由 UI 只提供合法选项
的机制保证——见断言：VA 下拉不含 Vb。）

自动化覆盖：切 Wiring→2E3W1P，Modbus 回读接线寄存器 + 全 24 通道 VA（奇 Va/偶 Vc）；
页面侧确认 VA 可选项={Va,Vc}（不含 Vb）、User 表 12 行、UC1 Phase A/C 可编辑、Phase B 固定。
说明：逐 User×Phase 遍历配置（用例表 step4~11）因当前设备 24 路输入通道已被前 8 个
User Channel 占满、Phase 下拉候选被耗尽，暂不纳入自动化，保留手工。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_MODE = "2 Element 3 Wire 1 Phase"


def test_001_07_case39(app_page, restore_wiring):
    nav_to_user_and_ct(app_page)
    switch_wiring(app_page, _MODE)

    # 1. 接线方式已生效
    verify_wiring_reg(_MODE)

    # 2. VA 规律：奇数通道 Va、偶数通道 Vc（页面抽样 + 全通道寄存器）
    assert read_va_cell_text(app_page, 0) == "Va", "row1 VA 应为 Va"
    assert read_va_cell_text(app_page, 1) == "Vc", "row2 VA 应为 Vc"
    verify_va_pattern(_MODE)

    # 3. VA 可选项 = {Va, Vc}，不含 Vb（非法相被 UI 挡住 → 对应 step3 保存失败语义）
    opts1 = set(open_va_options(app_page, 0))
    assert opts1 == {"Va", "Vc"}, f"row1 VA 可选项应为 {{Va,Vc}}，实际 {opts1}"
    opts2 = set(open_va_options(app_page, 1))
    assert opts2 == {"Va", "Vc"}, f"row2 VA 可选项应为 {{Va,Vc}}，实际 {opts2}"

    # 4. User and Channel Mapping 显示且 12 行
    assert user_mapping_visible(app_page), "2E3W1P 下 User Channel 区块应显示"
    assert user_channel_row_count(app_page) == 12, "User Channel 应为 12 行"

    # 5. UC1：Phase A/C 可配置，Phase B 固定(disabled)
    assert not phase_is_disabled(app_page, 0, "A"), "UC1 Phase A 应可编辑"
    assert not phase_is_disabled(app_page, 0, "C"), "UC1 Phase C 应可编辑"
    assert phase_is_disabled(app_page, 0, "B"), "UC1 Phase B 应为固定/disabled"
