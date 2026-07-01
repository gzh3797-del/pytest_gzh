# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case14
类别：D — 输入校验（Log File Name Prefix 超长）
功能：Rapid Logger 填写超长 Prefix 并保存时，页面应拒绝保存并显示错误提示

【预置条件】
  1. conftest.py driver fixture 已完成网关 Web 登录

【测试步骤】
  1. 进入 Rapid Logger 配置页面，将 Enable 设为 True
  2. 将 Post Channel 设为 Channel 3（HTTP/HTTPS）
  3. Log File Format 选 JSON，Log File Length 选 10 minutes
  4. Log File Name Format 选 Time Interval Format
  5. Log File Name Prefix 填写超长字符串：meter2_logger1_12345678910
  6. Log Interval 选 1 minute
  7. 点击 Save 按钮
  8. 检查页面是否弹出错误提示信息

【预期结果】
  - 页面显示错误提示（超出最大字符数等提示语）
  - 配置不被保存，Prefix 保持原值
"""
from datalog_page import RapidLoggerPage
from helpers import check_save_error

CASE_ID     = "TestCase_AcuHMI_003_04_case14"
LONG_PREFIX = "meter2_logger1_12345678910"


def test_case(driver):
    rl_page = RapidLoggerPage(driver)
    rl_page.navigate_to_rapid_logger()
    rl_page._set_enable(True)
    rl_page._select_el_by_text(rl_page._POST_CHANNEL_SELECT, "Post Channel 3", "Post Channel")
    rl_page._select_el_by_text(rl_page._LOG_FILE_FORMAT_SELECT, "json", "Log File Format")
    rl_page._select_el_by_text(rl_page._LOG_FILE_LENGTH_SELECT, "10 minutes", "Log File Length")
    rl_page._set_log_file_name_format("Time Interval Format")
    rl_page._fill(rl_page._LOG_FILE_NAME_PREFIX_INPUT, LONG_PREFIX, "Log File Name Prefix (超长)")
    rl_page._select_el_by_text(rl_page._LOG_INTERVAL_SELECT, "1 minute", "Log Interval")
    rl_page._safe_click(rl_page._SAVE_BTN, "Save")
    error_msg = check_save_error(driver)
    assert error_msg, (
        f"[{CASE_ID}] 超长 prefix 保存后，页面未显示错误信息（prefix='{LONG_PREFIX}'）"
    )
