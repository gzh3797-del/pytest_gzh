# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_03_case13
类别：E — Log Interval 联动验证
功能：Logger3 中，切换 Log File Length 时，Log Interval 下拉选项应与规格一一对应

【预置条件】
  1. conftest.py driver fixture 已完成网关 Web 登录

【测试步骤】
  依次将 Logger3 的 Log File Length 设为以下各值，每次读取 Log Interval 下拉选项并与期望值比对：

  | Log File Length | 期望 Log Interval 选项                                                           |
  |-----------------|---------------------------------------------------------------------------------|
  | 1 minute        | 1 minute                                                                        |
  | 5 minute        | 1 minute, 5 minute                                                              |
  | 10 minute       | 1 minute, 5 minute, 10 minute                                                   |
  | 15 minute       | 1 minute, 5 minute, 10 minute, 15 minute                                        |
  | 30 minute       | 1 minute, 5 minute, 10 minute, 15 minute, 30 minute                             |
  | 1 hour          | 1 minute, 5 minute, 10 minute, 15 minute, 30 minute, 1 hour                     |
  | 6 hour          | 1 minute, 5 minute, 10 minute, 15 minute, 30 minute, 1 hour, 6 hour             |
  | 12 hour         | 1 minute, 5 minute, 10 minute, 15 minute, 30 minute, 1 hour, 6 hour, 12 hour   |
  | 24 hour         | 5 minute, 10 minute, 15 minute, 30 minute, 1 hour, 6 hour, 12 hour, 1 day      |
  | 7 day           | 15 minute, 30 minute, 1 hour, 6 hour, 12 hour, 1 day, 7 day                    |
  | 1 month         | 1 hour, 6 hour, 12 hour, 1 day, 7 day, 1 month                                 |

【预期结果】
  - 每个 Log File Length 值对应的 Log Interval 下拉选项与上表完全一致（无缺失、无多余）
"""
from datalog_page import DataLoggerPage
from helpers import get_interval_options, EXPECTED_INTERVALS

CASE_ID  = "TestCase_AcuHMI_003_03_case13"
LOGGER_N = 3


def test_case(driver):
    dl_page = DataLoggerPage(driver)
    mismatches = []
    for length_text, expected in EXPECTED_INTERVALS.items():
        actual = get_interval_options(dl_page, LOGGER_N, length_text)
        if not actual:
            mismatches.append(f"  [{length_text}] 无法读取 LogInterval 下拉选项")
            continue
        missing = set(expected) - set(actual)
        extra   = set(actual) - set(expected)
        if missing or extra:
            mismatches.append(f"  [{length_text}] 缺失：{missing}  多余：{extra}")
    assert not mismatches, (
        f"[{CASE_ID}] Logger{LOGGER_N} 联动验证失败：\n" + "\n".join(mismatches)
    )
