# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_03_case11
类别：D — 输入校验（Log File Name Prefix 超长）
功能：Logger3 填写超长 Prefix 并保存时，页面应拒绝保存并显示错误提示

【预置条件】
  1. conftest.py driver fixture 已完成网关 Web 登录

【测试步骤】
  1. 进入 Data Loggers 页面，导航到 Logger3 配置
  2. 将 Enable 设为 True
  3. 在 Log File Name Prefix 输入框中填入超长字符串：
       meter2_logger1_12345678910（超出页面允许的最大字符数）
  4. 点击 Save 按钮
  5. 检查页面是否弹出错误提示信息

【预期结果】
  - 页面显示错误提示（如"超出最大字符数"等提示语）
  - 配置不被保存，Logger3 Prefix 保持原值
"""
from datalog_page import DataLoggerPage
from helpers import check_save_error

CASE_ID     = "TestCase_AcuHMI_003_03_case11"
LOGGER_N    = 3
LONG_PREFIX = "meter2_logger1_12345678910"


def test_case(driver):
    dl_page = DataLoggerPage(driver)
    dl_page.navigate_to_logger(LOGGER_N)
    dl_page._set_enable(True)
    dl_page._fill(dl_page._LOG_FILE_NAME_PREFIX_INPUT, LONG_PREFIX, "Log File Name Prefix (超长)")
    dl_page._safe_click(dl_page._SAVE_BTN, "Save")
    error_msg = check_save_error(driver)
    assert error_msg, (
        f"[{CASE_ID}] 超长 prefix 保存后，页面未显示错误信息（prefix='{LONG_PREFIX}'）"
    )
