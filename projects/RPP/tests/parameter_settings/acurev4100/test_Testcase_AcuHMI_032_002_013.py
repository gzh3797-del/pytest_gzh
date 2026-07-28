# -*- coding: utf-8 -*-
"""
TestCase_AcuHMI_032_002_013
模块：接入设备参数设置 → AcuRev4100（User Channel 命名）
标题：相关界面展示新名称且 THD 以及 Harmonic 页面除外

预期：改 UC5 Description=name5 后，涉及 User Channel 的界面展示 name5；
      Metering → THD / Harmonics 页面按现有设计不切换为 name5。
验证：
  - 正例：Metering 下 Realtime/Demand/Energy/Sequence，Meter Point 下拉选 UC5 均显示 name5；
  - 反例：Metering → THD/Harmonics 为 Input Channel 维度，其下拉文本不含 name5。
说明：实测这两页压根不按 User Channel 组织（选择器是 Input Channel N），故不显示 UC 名，
      与用例"THD/Harmonic 除外"吻合。PQ Event Log 列需真实 UC5 事件才可验，暂不纳入。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from _src_user_and_ct import *  # noqa: F401, F403

_NAME = "name5"


def test_032_002_013(app_page):
    # 改 UC5 并保存
    nav_to_user_and_ct(app_page)
    ensure_user_mapping(app_page)
    set_description(app_page, 4, _NAME)
    warning, msg = save_user_and_ct(app_page)
    assert not warning, f"UC5 改名应保存成功，实得警告：{msg!r}"
    verify_description_modbus(5, _NAME)

    # 正例：这些页面应展示 name5
    for leaf in METERING_UC_PAGES:
        nav_to_metering(app_page, leaf)
        shown = select_meter_point_uc(app_page, 5)
        assert _NAME in shown, f"Metering→{leaf} 应显示 {_NAME!r}，实得 {shown!r}"

    # 反例：THD / Harmonics 不展示 User Channel 名（Input Channel 维度）
    for leaf in METERING_INPUT_PAGES:
        nav_to_metering(app_page, leaf)
        txt = read_input_channel_select_text(app_page)
        assert _NAME not in txt, f"Metering→{leaf} 不应显示 User Channel 名，实得 {txt!r}"
